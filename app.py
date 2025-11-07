"""
app.py — AI класифікатор заявок із Gemini 2.5 Flash
Версія з повноцінним Resume Mode + автоочищення + автоархівування старих _out.xlsx
"""

# === 0. Автоматичне встановлення залежностей ===
import importlib, subprocess, sys
def ensure(pkg):
    try:
        importlib.import_module(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

for pkg in ["streamlit", "pandas", "openpyxl", "google-generativeai"]:
    ensure(pkg)

# === 1. Імпорти ===
import os, time, random, unicodedata, datetime
from pathlib import Path
import pandas as pd
import streamlit as st
import google.generativeai as genai
import gc

# === 2. Налаштування сторінки ===
st.set_page_config(page_title="AI Класифікатор (Gemini)", layout="wide")
st.title("🤖 AI Класифікатор заявок")

# Модель за замовчуванням (може бути перевизначена через config.txt)
MODEL_NAME = "gemini-2.5-flash-lite"


# === 3. Завантаження конфігурації ===
CONFIG_FILE = "config.txt"

# --- універсальна функція ---
def load_config():
    """
    Завантажує конфігурацію з двох джерел:
    1. Якщо є Streamlit Secrets (Cloud) → ключі GEMINI
    2. Якщо є локальний config.txt → модель, промпт, ключі (якщо локально)
    """
    cfg = {}

    # 1. Якщо є Streamlit Secrets (Cloud)
    try:    
        if hasattr(st, "secrets") and len(st.secrets) > 0:
            for key, value in st.secrets.items():
                # ✅ Якщо значення виглядає як список (наприклад "['a','b']") — перетворюємо
                try:
                    val = ast.literal_eval(str(value))
                except Exception:
                    val = str(value).strip()
                cfg[key.strip()] = val
    except Exception:
        pass # 🔹 ігноруємо помилку відсутності secrets.toml

    # 2. Якщо є локальний config.txt
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            lines = [line.strip() for line in f if line.strip() and not line.startswith("#")]
        for line in lines:
            if "=" in line:
                k, v = line.split("=", 1)
                cfg[k.strip()] = v.strip()

    return cfg


# --- створення прикладу config.txt, якщо відсутній ---
if not os.path.exists(CONFIG_FILE):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write("""# config.txt — приклад
MODEL_NAME = gemini-2.5-flash-lite

# Промпт (нижче текст шаблону)
PROMPT:
Визнач, який пункт з класифікатора (ID) найточніше відповідає опису проблеми.
Поверни лише один рядок у форматі:
ID=<id>
""")

st.sidebar.markdown(f"⚙️ Конфігурація: `{CONFIG_FILE}`")

# --- читаємо весь файл ---
content = ""
if os.path.exists(CONFIG_FILE):
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        content = f.read()

cfg = load_config()
st.sidebar.write("🧩 DEBUG: Secrets keys loaded →", list(st.secrets.keys()) if hasattr(st, "secrets") else "No secrets")
st.sidebar.write("🧩 DEBUG: cfg =", cfg)

# --- модель ---
MODEL_NAME = cfg.get("MODEL_NAME", "gemini-2.5-flash-lite")

# --- ключі ---
KEYS = []

# (1) якщо в Secrets є список GEMINI_KEYS
if "GEMINI_KEYS" in cfg and isinstance(cfg["GEMINI_KEYS"], (list, tuple)):
    KEYS = list(cfg["GEMINI_KEYS"])

# (2) якщо в Secrets один ключ GEMINI_KEY
elif "GEMINI_KEY" in cfg:
    KEYS = [cfg["GEMINI_KEY"]]

# (3) якщо локально є секція KEYS: у config.txt
elif "KEYS:" in content:
    keys_part = content.split("KEYS:")[1]
    keys_section = keys_part.split("PROMPT:")[0] if "PROMPT:" in keys_part else keys_part
    KEYS = [line.strip() for line in keys_section.splitlines() if line.strip() and not line.startswith("#")]


# --- промпт ---
if "PROMPT:" in content:
    prompt_template = content.split("PROMPT:")[1].strip()
else:
    prompt_template = "Визнач, який пункт з класифікатора (ID) найточніше відповідає опису проблеми.\nID=<id>"

# --- перевірка ---
if not KEYS:
    st.error("❌ Не знайдено жодного Gemini API ключа (ані в Streamlit Secrets, ані у config.txt).")
    st.stop()

# --- ініціалізація ключа ---
if "key_index" not in st.session_state:
    st.session_state.key_index = 0

def switch_key():
    st.session_state.key_index = (st.session_state.key_index + 1) % len(KEYS)
    genai.configure(api_key=KEYS[st.session_state.key_index])
    st.sidebar.info(f"🔄 Перемкнулися на ключ #{st.session_state.key_index+1}")

genai.configure(api_key=KEYS[st.session_state.key_index])
st.sidebar.write(f"🔹 Модель: `{MODEL_NAME}`")
st.sidebar.write(f"🔑 Активний ключ #{st.session_state.key_index+1} з {len(KEYS)}")

# === Відновлення попереднього стану з пам'яті (якщо сторінка перезавантажилась) ===
if "resume_df" in st.session_state:
    st.info("🔁 Відновлено незбережений прогрес із попередньої сесії.")
    resume_df = st.session_state["resume_df"]
    if "resume_path" in st.session_state:
        resume_path = st.session_state["resume_path"]

# === 4. Інтерфейс вибору файлів ===
st.header("Крок 1 — Вибір файлів")
col1, col2 = st.columns(2)
with col1:
    klass_file = st.file_uploader("📘 Вибери файл класифікатора (xlsx)", type=["xlsx"])
with col2:
    data_file = st.file_uploader("📗 Вибери файл даних (xlsx)", type=["xlsx"])

if not klass_file or not data_file:
    st.stop()

# === 5. Зчитування та очищення даних ===
def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Очищує всі клітинки від пробілів та порожніх рядків."""
    df = df.astype(str).applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df = df.replace({"": pd.NA, " ": pd.NA})
    df = df.fillna("")
    return df

klass_df = pd.read_excel(klass_file, dtype=str).fillna("")
klass_df = clean_dataframe(klass_df)

data_df = pd.read_excel(data_file, dtype=str).fillna("")
data_df = clean_dataframe(data_df)


# === 5.1. Збереження / відновлення налаштувань ===
config_path = Path(f"{Path(data_file.name).stem}_config.txt")

saved_settings = {}
if config_path.exists():
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            for line in f:
                if "=" in line:
                    key, value = line.strip().split("=", 1)
                    saved_settings[key] = value
        st.sidebar.success(f"⚙️ Завантажено попередні налаштування з {config_path.name}")
    except Exception as e:
        st.sidebar.warning(f"⚠️ Не вдалося прочитати {config_path.name}: {e}")

def save_current_settings():
    """Зберігає вибрані налаштування у <data_file>_config.txt"""
    cfg = {
        "klass_name_col": klass_name_col,
        "klass_id_col": klass_id_col,
        "klass_context_cols": ";".join(klass_context_cols),
        "data_text_cols": ";".join(data_text_cols),
        "out_name_col": out_name_col,
        "out_id_col": out_id_col
    }
    with open(config_path, "w", encoding="utf-8") as f:
        for k, v in cfg.items():
            f.write(f"{k}={v}\n")
    st.sidebar.info(f"💾 Збережено налаштування у {config_path.name}")



st.success(f"✅ Завантажено: класифікатор ({len(klass_df)}), дані ({len(data_df)})")

# === 6. Вибір колонок ===
st.header("Крок 2 — Вибір колонок")

klass_name_col = st.selectbox(
    "🔹 Колонка з Klassificator (назва):",
    list(klass_df.columns),
    index=(
        list(klass_df.columns).index(saved_settings.get("klass_name_col"))
        if saved_settings.get("klass_name_col") in klass_df.columns
        else 0
    )
)

klass_id_col = st.selectbox(
    "🔹 Колонка з Klassificator (ID):",
    list(klass_df.columns),
    index=(
        list(klass_df.columns).index(saved_settings.get("klass_id_col"))
        if saved_settings.get("klass_id_col") in klass_df.columns
        else 0
    )
)

klass_context_cols = st.multiselect(
    "📊 Колонки контексту (Klassificator):",
    [c for c in klass_df.columns if c not in [klass_name_col, klass_id_col]],
    default=[
        c for c in saved_settings.get("klass_context_cols", "").split(";")
        if c in klass_df.columns
    ]
)

data_text_cols = st.multiselect(
    "🧩 Колонки для аналізу (файл даних):",
    list(data_df.columns),
    default=[
        c for c in saved_settings.get("data_text_cols", "").split(";")
        if c in data_df.columns
    ]
)

out_name_col = st.selectbox(
    "💾 Колонка для результату (назва):",
    list(data_df.columns),
    index=(
        list(data_df.columns).index(saved_settings.get("out_name_col"))
        if saved_settings.get("out_name_col") in data_df.columns
        else 0
    )
)

out_id_col = st.selectbox(
    "💾 Колонка для результату (ID):",
    list(data_df.columns),
    index=(
        list(data_df.columns).index(saved_settings.get("out_id_col"))
        if saved_settings.get("out_id_col") in data_df.columns
        else 0
    )
)


# === 7. Промпт із config.txt ===
prompt_template = prompt_template.strip()

# === 8. Підготовка до Resume ===
out_path = Path(f"{Path(data_file.name).stem}_out.xlsx")

resume_mode = False
run_fresh = False

if out_path.exists():
    existing = pd.read_excel(out_path, dtype=str).fillna("")
    done_rows = existing[out_id_col].astype(str).str.strip() != ""
    done_count = done_rows.sum()
    total_rows = len(existing)

    st.warning(f"📄 Знайдено попередній результат: `{out_path.name}` ({done_count}/{total_rows} оброблено)")
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        run_fresh = st.button("▶ Почати спочатку (очистити)")
    with col_r2:
        resume_mode = st.button(f"🔁 Продовжити ({total_rows - done_count} рядків лишилось)")

    # якщо користувач натиснув "Почати спочатку" — перейменовуємо старий файл
    if run_fresh:
        timestamp = datetime.datetime.now().strftime("%y%m%d-%H%M")
        archived_path = out_path.with_name(out_path.stem + f"_{timestamp}.xlsx")
        try:
            os.rename(out_path, archived_path)
            st.info(f"📦 Старий результат перейменовано у `{archived_path.name}`")
        except Exception as e:
            st.warning(f"⚠ Не вдалося перейменувати старий файл: {e}")
else:
    run_fresh = st.button("▶ Почати класифікацію")

if not (run_fresh or resume_mode):
    st.stop()

# === 9. Формування списку кандидатів ===
def build_candidates(df):
    lines = []
    for _, row in df.iterrows():
        rid = str(row.get(klass_id_col, "")).strip()
        name = str(row.get(klass_name_col, "")).strip()
        context = " ".join(str(row[c]) for c in klass_context_cols)
        lines.append(f"{rid} | {name} | {context}")
    return "\n".join(lines)

candidates_text = build_candidates(klass_df)

# === 10. Логіка обробки ===
model = genai.GenerativeModel(MODEL_NAME)

# Якщо режим Resume — читаємо існуючий файл
if resume_mode and out_path.exists():
    result_df = existing.copy()
    rows_to_process = result_df.index[result_df[out_id_col].astype(str).str.strip() == ""].tolist()
    st.info(f"🔄 Продовження: залишилось {len(rows_to_process)} рядків")
else:
    result_df = data_df.copy()
    rows_to_process = list(result_df.index)
    # === Зберігаємо початковий стан у сесію (щоб можна було відновити після перезавантаження) ===
    st.session_state["resume_df"] = result_df
    st.session_state["resume_path"] = str(out_path)
    if out_path.exists():
        timestamp = datetime.datetime.now().strftime("%y%m%d-%H%M")
        os.rename(out_path, f"{out_path.stem}_{timestamp}.xlsx")

total = len(rows_to_process)
st.info(f"🔄 Для обробки: {total} рядків")

progress = st.progress(0)
status = st.empty()

# Зберігаємо поточні налаштування перед запуском
if run_fresh or resume_mode:
    save_current_settings()

# цикл for
for i, idx in enumerate(rows_to_process, start=1):
    row = result_df.loc[idx]
    text = " ".join(str(row[c]) for c in data_text_cols)

    prompt = f"""{prompt_template}

Опис проблеми:
{text}

Список варіантів:
{candidates_text}
"""

    # --- нова логіка з повторними спробами ---
    for attempt in range(3):  # максимум 3 спроби на одну заявку
        try:
            resp = model.generate_content(prompt)
            txt = resp.text.strip()
            if "ID=" in txt:
                txt = txt.split("ID=")[-1].strip()
            txt = txt.replace("\n", "").replace(";", "").strip()

            # Знайдемо назву за ID
            match = klass_df[klass_df[klass_id_col].astype(str).str.strip() == txt]
            name_val = match[klass_name_col].iloc[0] if not match.empty else "НЕ ЗНАЙДЕНО"

            result_df.at[idx, out_id_col] = txt
            result_df.at[idx, out_name_col] = name_val
            break  # 🟢 успішно, виходимо з циклу повторних спроб

        except Exception as e:
            err = str(e)
            if "429" in err:
                st.sidebar.warning(f"⚠️ Перевищено ліміт, перемикаємо ключ (спроба {attempt+1}/3)...")
                switch_key()
                model = genai.GenerativeModel(MODEL_NAME)
                time.sleep(5)
                continue  # 🔁 повторити ту ж заявку ще раз
            else:
                result_df.at[idx, out_id_col] = ""
                result_df.at[idx, out_name_col] = f"ERROR: {err}"
                break  # ❌ інша помилка — виходимо з повторних спроб

    # --- Періодичне збереження ---
    if i % 2 == 0 or i == total:
        result_df.to_excel(out_path, index=False)
        # Оновлення копії в пам’яті Streamlit
        st.session_state["resume_df"] = result_df
        progress.progress(i / total)
        status.markdown(f"Оброблено: **{i} / {total}** ({time.strftime('%H:%M:%S')})")
        time.sleep(0.5)
        # збереження прогресу
        # save_progress()  # збереження вже виконується через session_state
       
    # новий блок очищення пам’яті кожні 100 рядків
    if i % 100 == 0:
        gc.collect()

# === 11. Завершення обробки ===
result_df.to_excel(out_path, index=False)
# Оновлення копії в пам’яті Streamlit
st.session_state["resume_df"] = result_df

st.success("✅ Обробку завершено успішно!")
st.balloons()

# --- Створення буфера у пам’яті ---
from io import BytesIO
buffer = BytesIO()
result_df.to_excel(buffer, index=False)
buffer.seek(0)

# --- Пропозиція завантажити результат ---
st.download_button(
    label="⬇️ Завантажити результат (Excel)",
    data=buffer.getvalue(),
    file_name=f"{Path(data_file.name).stem}_out.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# --- Інформаційне повідомлення ---
if os.access(".", os.W_OK):
    st.info(f"📁 Результат також збережено локально у `{out_path.name}`")
else:
    st.info("☁️ Файл створено у пам'яті (RAM). У Streamlit Cloud збереження локально не підтримується.")

