import streamlit as st
import random
import pandas as pd
import altair as alt
from datetime import datetime, timedelta, date
import json
import html
import hmac
from pathlib import Path
import os

st.set_page_config(
    page_title="Gestão de Reservas",
    page_icon="🌅",
    layout="wide",
)

SITE_PASSWORD = st.secrets.get("SITE_PASSWORD", "Almograve2026")
AUTH_ENABLED = False


def require_password():
    if not AUTH_ENABLED or st.session_state.get("is_authenticated", False):
        return

    st.title("Acesso protegido")
    st.write("Insere a palavra-passe para entrar na aplicação.")

    with st.form("login_form"):
        typed_password = st.text_input("Palavra-passe", type="password")
        login_submit = st.form_submit_button("Entrar")

    if login_submit:
        if hmac.compare_digest(str(typed_password), str(SITE_PASSWORD)):
            st.session_state["is_authenticated"] = True
            st.rerun()
        else:
            st.error("Palavra-passe incorreta.")

    st.stop()


require_password()

# Painel central de cores: muda apenas aqui para atualizar todo o visual.

THEME = {
    "primary": "#008080",
    "primary_dark": "#006666",
    "tab_active": "#ff7f7f",
    "menu_hover_bg": "#e6f7f7",
    "alert_bg": "#ffe5e5",
    "alert_border": "#ff7f7f",
    "alert_text": "#8a2b2b",
    "chart_ok": "#00d084",
    "chart_over": "#e74c3c",
}

# Estrutura fixa do checklist de saídas (exemplo, ajuste conforme necessário)
CHECKLIST_STRUCTURE = [
    ("Alojamento 1", ["Quarto 1", "Quarto 2"]),
    ("Alojamento 2", ["Quarto 3", "Quarto 4"]),
]

ALOJAMENTO_BUTTON_COLORS = {
    "Alojamento 1": "#ef4444",
    "Alojamento 2": "#3b82f6",
}

st.markdown(
        f"""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;700&display=swap');
            @import url('https://fonts.googleapis.com/icon?family=Material+Icons');

            :root {{
                --primary-color: {THEME['primary']};
                color-scheme: light;
            }}

            .stApp {{ background-color: #f8f9fa; }}

            /* Força legibilidade em modo dark/system no telemóvel */
            html[data-theme="dark"] .stApp,
            html[data-theme="dark"] [data-testid="stAppViewContainer"],
            html[data-theme="dark"] [data-testid="stHeader"] {{
                background-color: #f8f9fa !important;
            }}

            html[data-theme="dark"] [data-testid="stAppViewContainer"] * {{
                color: #1f2933 !important;
            }}

            html[data-theme="dark"] [data-baseweb="input"] input,
            html[data-theme="dark"] [data-baseweb="select"] > div,
            html[data-theme="dark"] [data-testid="stDataFrame"] {{
                background: #ffffff !important;
                color: #1f2933 !important;
            }}

            html, body, [class*="st-"] {{
                font-family: 'Inter', sans-serif;
            }}

            button[data-baseweb="tab"] p {{
                font-size: 1.1rem;
                font-weight: 700;
            }}

            button[data-baseweb="tab"][aria-selected="true"] p {{
                color: {THEME['tab_active']} !important;
            }}

            button[data-baseweb="tab"][aria-selected="true"] {{
                border-bottom-color: {THEME['tab_active']} !important;
            }}

            /* Hover/focus em tom verde-azulado em vez de vermelho */
            .stButton > button:hover,
            .stDownloadButton > button:hover {{
                border-color: {THEME['primary']} !important;
                color: {THEME['primary_dark']} !important;
            }}

            [data-baseweb="select"]:focus-within,
            [data-baseweb="input"]:focus-within {{
                box-shadow: 0 0 0 1px {THEME['primary']} !important;
                border-color: {THEME['primary']} !important;
            }}

            div[role="listbox"] ul li:hover,
            div[role="option"]:hover {{
                background-color: {THEME['menu_hover_bg']} !important;
            }}

            /* Arredondar botões e outros elementos */
            .stButton > button,
            .stDownloadButton > button,
            button[data-baseweb="button"] {{
                border-radius: 15px !important;
            }}

            /* No telemóvel, remove o menu flutuante (olho/download/fullscreen) */
            @media (max-width: 900px) {{
                [data-testid="stElementToolbar"] {{
                    display: none !important;
                    visibility: hidden !important;
                }}
            }}

            /* Esconde o input nativo do browser no file uploader para evitar texto duplicado */
            [data-testid="stFileUploader"] input[type="file"] {{
                opacity: 0 !important;
                position: absolute !important;
                width: 0 !important;
                height: 0 !important;
                overflow: hidden !important;
                pointer-events: none !important;
            }}

        </style>
        """,
        unsafe_allow_html=True,
)

APP_DIR = Path(__file__).resolve().parent
RESERVAS_FILE = APP_DIR / "reservas.json"
QUARTOS_FILE = APP_DIR / "quartos_disponiveis.json"
NOTAS_GERAIS_FILE = APP_DIR / "notas_gerais.json"
SAIDAS_FILE = APP_DIR / "saidas_checklist.json"
TRANSFERS_FILE = APP_DIR / "transfers.json"
RESERVA_KEY_COLS = ["Nome", "Check-in", "Check-out", "Pessoas", "Unidade", "Alojamento"]
IMPORT_MATCH_COLS = ["Nome", "Check-in", "Alojamento"]
IMPORT_UPDATE_COLS = ["Check-out", "Pessoas", "Unidade"]
DISPLAY_COL_ORDER = [
    "Alojamento",
    "Nome",
    "Check-in",
    "Check-out",
    "Pessoas",
    "Unidade",
    "Origem",
    "Hora PA",
    "PA pago",
    "Notas",
]

ALOJAMENTOS = ["ABH", "AFH", "PIPO", "DUNAS", "DUNAS2", "FOZ", "ESCAPE", "MARES"]


def normalize_alojamento(value):
    import re

    if pd.isna(value):
        return ""

    txt = str(value).strip().upper()
    if not txt:
        return ""

    # Prioriza códigos mais longos para evitar DUNAS2 cair em DUNAS.
    for code in sorted(ALOJAMENTOS, key=len, reverse=True):
        if re.search(rf"\b{re.escape(code)}\b", txt):
            return code

    # Remove prefixos visuais (emoji/símbolos) e espaços extra.
    txt = re.sub(r"^[^A-Z0-9]+", "", txt)
    return txt


def suggest_alojamento_from_filename(filename):
    suggested = normalize_alojamento(filename)
    if suggested in ALOJAMENTOS:
        return suggested
    return ALOJAMENTOS[0]


def expand_unidade_sequence(base_unidade, quantidade):
    import re

    quantidade = max(1, int(quantidade))
    base_txt = str(base_unidade).strip()
    if quantidade == 1 or not base_txt:
        return [base_txt]

    m = re.match(r"^(quarto|cama)\s*(\d+)$", base_txt, flags=re.IGNORECASE)
    if not m:
        return [base_txt]

    prefix = m.group(1).capitalize()
    start = int(m.group(2))
    return [f"{prefix} {n}" for n in range(start, start + quantidade)]

# ── Supabase (sincronização na nuvem, opcional) ──────────────────────────────
USE_SUPABASE = False
_supabase_client = None

try:
    from supabase import create_client
    _url = st.secrets.get("SUPABASE_URL", "")
    _key = st.secrets.get("SUPABASE_KEY", "")
    if _url and _key:
        _supabase_client = create_client(_url, _key)
        USE_SUPABASE = True
except Exception:
    pass


def _serialize_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def show_pink_alert(message):
    safe_message = html.escape(str(message))
    st.markdown(
        f"""
        <div style=\"background:{THEME['alert_bg']};border-left:5px solid {THEME['alert_border']};color:{THEME['alert_text']};padding:0.85rem 1rem;border-radius:0.5rem;margin:0.35rem 0;\">{safe_message}</div>
        """,
        unsafe_allow_html=True,
    )


def _json_candidates_for(filename):
    """Prioritize the app folder, then accept legacy files saved from another cwd."""
    candidates = [APP_DIR / filename]
    cwd_candidate = Path.cwd() / filename
    try:
        if cwd_candidate.resolve() != (APP_DIR / filename).resolve():
            candidates.append(cwd_candidate)
    except Exception:
        if str(cwd_candidate) != str(APP_DIR / filename):
            candidates.append(cwd_candidate)
    return candidates


def _atomic_write_json(path, payload):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    with tmp_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    tmp_path.replace(path)


def _parse_reservas_payload(data):
    if isinstance(data, dict):
        last_saved_at = data.get("last_saved_at")
        reservas = data.get("reservas", [])
        if isinstance(reservas, list):
            return pd.DataFrame(reservas), last_saved_at
    if isinstance(data, list):
        return pd.DataFrame(data), None
    raise ValueError("Formato de reservas inválido")


def load_reservas():
    if USE_SUPABASE:
        try:
            resp = _supabase_client.table("reservas").select("data").eq("id", 1).execute()
            if resp.data:
                data = resp.data[0]["data"]
                if isinstance(data, dict):
                    last_saved_at = data.get("last_saved_at")
                    if last_saved_at:
                        parsed_last_saved_at = pd.to_datetime(last_saved_at, errors="coerce")
                        if pd.notna(parsed_last_saved_at):
                            st.session_state["last_saved_at"] = parsed_last_saved_at.to_pydatetime()
                    reservas = data.get("reservas", [])
                    if isinstance(reservas, list):
                        return pd.DataFrame(reservas)
                if isinstance(data, list):
                    migrated_at = datetime.now()
                    st.session_state["last_saved_at"] = migrated_at
                    try:
                        _supabase_client.table("reservas").upsert(
                            {
                                "id": 1,
                                "data": {
                                    "reservas": data,
                                    "last_saved_at": migrated_at.isoformat(),
                                },
                            }
                        ).execute()
                    except Exception:
                        pass
                    return pd.DataFrame(data)
        except Exception as e:
            show_pink_alert(f"Erro ao carregar do Supabase: {e}")
        return pd.DataFrame()

    parse_errors = []
    for candidate in _json_candidates_for("reservas.json"):
        if not candidate.exists():
            continue

        try:
            with candidate.open("r", encoding="utf-8") as f:
                data = json.load(f)

            df_loaded, last_saved_at = _parse_reservas_payload(data)
            if last_saved_at:
                parsed_last_saved_at = pd.to_datetime(last_saved_at, errors="coerce")
                if pd.notna(parsed_last_saved_at):
                    st.session_state["last_saved_at"] = parsed_last_saved_at.to_pydatetime()

            # Migra automaticamente para o ficheiro principal da app.
            if candidate != RESERVAS_FILE:
                save_reservas(df_loaded)
                show_pink_alert(
                    "Reservas recuperadas de um ficheiro antigo e migradas para a pasta da aplicação."
                )

            return df_loaded
        except Exception as e:
            parse_errors.append(f"{candidate}: {e}")

    if parse_errors:
        show_pink_alert(
            "Erro ao carregar reservas guardadas: " + " | ".join(parse_errors)
        )

    return pd.DataFrame()


def save_reservas(df):
    records = []
    for record in df.to_dict(orient="records"):
        clean_record = {k: _serialize_value(v) for k, v in record.items()}
        records.append(clean_record)

    saved_at = datetime.now()
    payload = {
        "reservas": records,
        "last_saved_at": saved_at.isoformat(),
    }

    if USE_SUPABASE:
        try:
            _supabase_client.table("reservas").upsert({"id": 1, "data": payload}).execute()
            st.session_state["last_saved_at"] = saved_at
            return
        except Exception as e:
            show_pink_alert(f"Erro ao guardar no Supabase: {e}")

    _atomic_write_json(RESERVAS_FILE, payload)
    st.session_state["last_saved_at"] = saved_at


def load_quartos():
    parse_errors = []
    for candidate in _json_candidates_for("quartos_disponiveis.json"):
        if not candidate.exists():
            continue
        try:
            with candidate.open("r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                if candidate != QUARTOS_FILE:
                    save_quartos(data)
                return data
        except Exception as e:
            parse_errors.append(f"{candidate}: {e}")

    if parse_errors:
        show_pink_alert(
            "Erro ao carregar quartos guardados: " + " | ".join(parse_errors)
        )
    return []


def save_quartos(quartos_list):
    _atomic_write_json(QUARTOS_FILE, quartos_list)


def load_notas_gerais():
    parse_errors = []
    for candidate in _json_candidates_for("notas_gerais.json"):
        if not candidate.exists():
            continue
        try:
            with candidate.open("r", encoding="utf-8") as f:
                data = json.load(f)

            if isinstance(data, dict):
                notas = data.get("notas_gerais", "")
            elif isinstance(data, str):
                notas = data
            else:
                notas = ""

            notas = "" if notas is None else str(notas)

            if candidate != NOTAS_GERAIS_FILE:
                save_notas_gerais(notas)

            return notas
        except Exception as e:
            parse_errors.append(f"{candidate}: {e}")

    if parse_errors:
        show_pink_alert(
            "Erro ao carregar notas gerais: " + " | ".join(parse_errors)
        )

    return ""


def save_notas_gerais(texto):
    payload = {
        "notas_gerais": "" if texto is None else str(texto).strip(),
        "last_saved_at": datetime.now().isoformat(),
    }
    _atomic_write_json(NOTAS_GERAIS_FILE, payload)


def load_transfers():
    if USE_SUPABASE:
        try:
            resp = _supabase_client.table("reservas").select("data").eq("id", 2).execute()
            if resp.data:
                data = resp.data[0]["data"]
                if isinstance(data, list):
                    return data
                if isinstance(data, dict):
                    return data.get("transfers", [])
        except Exception:
            pass
        return []
    if TRANSFERS_FILE.exists():
        try:
            with TRANSFERS_FILE.open("r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, list) else []
        except Exception:
            pass
    return []


def save_transfers(transfers):
    if USE_SUPABASE:
        try:
            _supabase_client.table("reservas").upsert({"id": 2, "data": transfers}).execute()
        except Exception as e:
            show_pink_alert(f"Erro ao guardar transfers: {e}")
        return
    _atomic_write_json(TRANSFERS_FILE, transfers)


def data_referencia_checklist(df=None):
    """Devolve a data de referência da checklist: a data de check-in mais frequente
    nos dados carregados. Se não houver dados, usa a data de hoje."""
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        return date.today().isoformat()
    try:
        checkins = pd.to_datetime(df["Check-in"], errors="coerce").dropna()
        if checkins.empty:
            return date.today().isoformat()
        return checkins.dt.date.mode()[0].isoformat()
    except Exception:
        return date.today().isoformat()


def load_saidas_checklist(df=None):
    """Carrega checklist guardada. Se for de outro período de trabalho, devolve {}."""
    if not SAIDAS_FILE.exists():
        return {}
    try:
        with SAIDAS_FILE.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            saved_date = data.get("_date")
            ref_date = data_referencia_checklist(df)
            if saved_date != ref_date:
                return {}
            return {k: bool(v) for k, v in data.items() if not k.startswith("_")}
    except Exception:
        pass
    return {}


def save_saidas_checklist(checklist: dict, df=None):
    payload = {k: bool(v) for k, v in checklist.items() if not k.startswith("_")}
    ref_date = data_referencia_checklist(df)
    if hasattr(ref_date, 'isoformat'):
        ref_date = ref_date.isoformat()
    payload["_date"] = ref_date
    _atomic_write_json(SAIDAS_FILE, payload)


def refresh_data_from_storage():
    refreshed_reservas = sanitize_optional_columns(load_reservas())
    if not refreshed_reservas.empty and "Origem" not in refreshed_reservas.columns:
        refreshed_reservas["Origem"] = "Importada"
    elif not refreshed_reservas.empty:
        refreshed_reservas["Origem"] = (
            refreshed_reservas["Origem"]
            .astype(str)
            .str.strip()
            .replace({"": "Importada", "none": "Importada", "nan": "Importada", "None": "Importada"})
        )

    refreshed_reservas = normalize_pessoas_column(refreshed_reservas)
    st.session_state["reservas_df"] = refreshed_reservas.copy()
    st.session_state["reservas_editor_df"] = refreshed_reservas.copy()
    st.session_state["quartos_disponiveis"] = load_quartos()
    st.session_state["notas_gerais_pa"] = load_notas_gerais()
    st.session_state["notas_gerais_pa_editor"] = st.session_state["notas_gerais_pa"]
    st.session_state["saidas_checklist"] = load_saidas_checklist(refreshed_reservas)
    st.session_state["transfers"] = load_transfers()


def get_last_saved_text():
    last = st.session_state.get("last_saved_at")
    if last:
        if isinstance(last, pd.Timestamp):
            last = last.to_pydatetime()
        return last.strftime("%d/%m/%Y %H:%M:%S")

    if not USE_SUPABASE:
        if not RESERVAS_FILE.exists():
            return "Ainda não guardado"
        ts = datetime.fromtimestamp(RESERVAS_FILE.stat().st_mtime)
        return ts.strftime("%d/%m/%Y %H:%M:%S")

    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def _normalize_key_value(value, column_name):
    if pd.isna(value):
        return ""

    if column_name in ("Check-in", "Check-out"):
        dt = pd.to_datetime(value, errors="coerce")
        if pd.isna(dt):
            return str(value).strip().lower()
        return dt.strftime("%Y-%m-%d")

    if column_name == "Pessoas":
        num = pd.to_numeric(value, errors="coerce")
        if pd.isna(num):
            return str(value).strip().lower()
        return str(int(num))

    return str(value).strip().lower()


def _build_reserva_key(row):
    return tuple(_normalize_key_value(row.get(col), col) for col in RESERVA_KEY_COLS)


def _build_import_match_key(row):
    return tuple(_normalize_key_value(row.get(col), col) for col in IMPORT_MATCH_COLS)


def merge_new_reservas(df_existing, df_incoming):
    if df_incoming.empty:
        return df_existing.copy(), 0

    if df_existing.empty:
        return df_incoming.copy(), len(df_incoming)

    # Garantir colunas para não perder campos já existentes (ex.: Hora PA, PA pago)
    all_cols = list(dict.fromkeys([*df_existing.columns.tolist(), *df_incoming.columns.tolist()]))
    df_existing = df_existing.reindex(columns=all_cols)
    df_incoming = df_incoming.reindex(columns=all_cols)

    existing_keys = set(df_existing.apply(_build_reserva_key, axis=1).tolist())
    incoming_keys = df_incoming.apply(_build_reserva_key, axis=1)
    mask_new = ~incoming_keys.isin(existing_keys)
    df_new = df_incoming[mask_new]

    if df_new.empty:
        return df_existing.copy(), 0

    merged = pd.concat([df_existing, df_new], ignore_index=True)
    return merged, len(df_new)


def merge_imported_reservas(df_existing, df_incoming):
    if df_incoming.empty:
        return df_existing.copy(), 0, 0

    incoming = df_incoming.copy()
    incoming["Origem"] = "Importada"

    if df_existing.empty:
        return incoming.copy(), len(incoming), 0

    all_cols = list(dict.fromkeys([*df_existing.columns.tolist(), *incoming.columns.tolist()]))
    existing = df_existing.reindex(columns=all_cols).copy()
    incoming = incoming.reindex(columns=all_cols).copy()

    existing_key_to_idx = {}
    for idx, row in existing.iterrows():
        key = _build_import_match_key(row)
        if key not in existing_key_to_idx:
            existing_key_to_idx[key] = idx

    added_count = 0
    updated_count = 0

    for _, in_row in incoming.iterrows():
        key = _build_import_match_key(in_row)
        match_idx = existing_key_to_idx.get(key)

        if match_idx is None:
            new_row_df = pd.DataFrame([in_row]).reindex(columns=all_cols)
            existing = pd.concat([existing, new_row_df], ignore_index=True)
            existing_key_to_idx[key] = len(existing) - 1
            added_count += 1
            continue

        row_changed = False
        for col in IMPORT_UPDATE_COLS:
            old_val = existing.at[match_idx, col] if col in existing.columns else None
            new_val = in_row.get(col)

            old_is_na = pd.isna(old_val)
            new_is_na = pd.isna(new_val)
            if old_is_na and new_is_na:
                continue

            if (old_is_na and not new_is_na) or (not old_is_na and new_is_na) or str(old_val).strip() != str(new_val).strip():
                existing.at[match_idx, col] = new_val
                row_changed = True

        if row_changed:
            updated_count += 1

    return existing, added_count, updated_count


def build_suggested_times():
    start_time = datetime.strptime("06:30", "%H:%M")
    end_time = datetime.strptime("10:00", "%H:%M")
    current_time = start_time
    suggested_times = []
    while current_time <= end_time:
        suggested_times.append(current_time.strftime("%H:%M"))
        current_time += timedelta(minutes=15)
    return suggested_times


def sanitize_optional_columns(df):
    if df.empty:
        return df

    cleaned = df.copy()
    for col in ("Hora PA", "PA pago", "Notas"):
        if col in cleaned.columns:
            cleaned[col] = cleaned[col].apply(
                lambda v: None
                if pd.isna(v) or str(v).strip().lower() in {"none", "nan", "nat", ""}
                else str(v).strip()
            )
    return cleaned


def build_occupation_data(df_pa, suggested_times):
    occupation_data = []
    for time_slot in suggested_times:
        time_obj = datetime.strptime(time_slot, "%H:%M")
        occupation = 0
        for _, person_row in df_pa.iterrows():
            person_time_str = person_row["Hora PA"]
            if person_time_str:
                person_time = datetime.strptime(person_time_str, "%H:%M")
                person_end = person_time + timedelta(minutes=45)
                if person_time <= time_obj < person_end:
                    pessoas_val = pd.to_numeric(person_row.get("Pessoas"), errors="coerce")
                    occupation += int(pessoas_val) if pd.notna(pessoas_val) else 0
        occupation_data.append({"Hora": time_slot, "Pessoas": occupation})
    return occupation_data


def build_overcrowding_messages(df, suggested_times, threshold=16):
    if df.empty or "Hora PA" not in df.columns:
        return []

    df_pa = df.copy()
    df_pa = df_pa[df_pa["Hora PA"].notna() & (df_pa["Hora PA"] != "")]
    if df_pa.empty:
        return []

    messages = []
    for row in build_occupation_data(df_pa, suggested_times):
        if row["Pessoas"] >= threshold:
            pessoas_num = pd.to_numeric(row.get("Pessoas"), errors="coerce")
            pessoas_total = int(pessoas_num) if pd.notna(pessoas_num) else 0
            messages.append(f"⚠ {row['Hora']} → {pessoas_total} pessoas")
    return messages


def format_checkin_checkout(checkin_value, checkout_value):
    checkin_dt = pd.to_datetime(checkin_value, errors="coerce")
    checkout_dt = pd.to_datetime(checkout_value, errors="coerce")

    if pd.notna(checkin_dt) and pd.notna(checkout_dt):
        return f"{checkin_dt.strftime('%d/%m')}-{checkout_dt.strftime('%d/%m')}"
    if pd.notna(checkin_dt):
        return checkin_dt.strftime("%d/%m")
    if pd.notna(checkout_dt):
        return checkout_dt.strftime("%d/%m")
    return ""


def total_pessoas_col(df):
    if df.empty or "Pessoas" not in df.columns:
        return 0
    return int(pd.to_numeric(df["Pessoas"], errors="coerce").fillna(0).sum())


def normalize_pessoas_value(value):
    import re

    if pd.isna(value):
        return 1

    numeric = pd.to_numeric(value, errors="coerce")
    if pd.notna(numeric):
        return max(1, int(numeric))

    txt = str(value).strip()
    match = re.search(r"\d+", txt)
    if match:
        return max(1, int(match.group(0)))

    return 1


def normalize_pessoas_column(df):
    if df.empty or "Pessoas" not in df.columns:
        return df

    normalized = df.copy()
    normalized["Pessoas"] = normalized["Pessoas"].apply(normalize_pessoas_value)
    return normalized


def total_hospedes_esta_noite(df, ref_date=None):
    return total_pessoas_col(df)


def _existing_date_options(df, column_name):
    if df.empty or column_name not in df.columns:
        return []

    parsed = pd.to_datetime(df[column_name], errors="coerce").dropna()
    if parsed.empty:
        return []

    unique_dates = sorted({d.date() for d in parsed.tolist()})
    return unique_dates


def _choose_date_with_suggestions(label, suggestions, default_date, key_prefix):
    use_manual = st.checkbox(
        f"{label} manual",
        value=not bool(suggestions),
        key=f"{key_prefix}_manual",
        help="Ativa para inserir a data manualmente.",
    )

    if suggestions and not use_manual:
        safe_default = default_date if default_date in suggestions else suggestions[0]
        return st.selectbox(
            label,
            options=suggestions,
            index=suggestions.index(safe_default),
            format_func=lambda d: d.strftime("%d/%m/%Y"),
            help="Datas sugeridas com base nos outros hóspedes.",
            key=f"{key_prefix}_suggested",
        )

    return st.date_input(
        label,
        value=default_date,
        key=f"{key_prefix}_manual_date",
    )


def _is_duplicate_reserva(df_existing, reserva_row, ignore_index=None):
    if df_existing.empty:
        return False

    target_key = _build_reserva_key(pd.Series(reserva_row))
    for idx, existing_row in df_existing.iterrows():
        if ignore_index is not None and idx == ignore_index:
            continue
        if _build_reserva_key(existing_row) == target_key:
            return True
    return False


def detect_conflicts(df):
    """Detecta reservas sobrepostas no mesmo alojamento+unidade.
    Retorna lista de strings descritivas dos conflitos encontrados.
    Ignora unidades vazias (reservas de propriedade inteira são tratadas separadamente).
    """
    if df is None or df.empty:
        return []

    conflicts = []

    def _parse_date(v):
        try:
            return pd.to_datetime(v).date()
        except Exception:
            return None

    rows = []
    for idx, row in df.iterrows():
        aloj = str(row.get("Alojamento", "") or "").strip()
        unidade = str(row.get("Unidade", "") or "").strip()
        checkin = _parse_date(row.get("Check-in"))
        checkout = _parse_date(row.get("Check-out"))
        nome = str(row.get("Nome", "") or "").strip()
        if aloj and unidade and checkin and checkout and checkin < checkout:
            rows.append((idx, aloj, unidade, checkin, checkout, nome))

    seen = set()
    for i, (idx_a, aloj_a, uni_a, ci_a, co_a, nome_a) in enumerate(rows):
        for idx_b, aloj_b, uni_b, ci_b, co_b, nome_b in rows[i + 1:]:
            if aloj_a != aloj_b or uni_a.lower() != uni_b.lower():
                continue
            # Sobreposição: [ci_a, co_a[ ∩ [ci_b, co_b[ ≠ ∅
            if ci_a < co_b and ci_b < co_a:
                key = tuple(sorted([idx_a, idx_b]))
                if key not in seen:
                    seen.add(key)
                    conflicts.append(
                        f"⚠ Conflito: **{aloj_a} — {uni_a}** reservado por **{nome_a}** ({ci_a.strftime('%d/%m')}–{co_a.strftime('%d/%m')}) e **{nome_b}** ({ci_b.strftime('%d/%m')}–{co_b.strftime('%d/%m')})"
                    )
    return conflicts


ALOJAMENTO_BADGES = {
    "ABH": "🟥",
    "AFH": "🟦",
    "PIPO": "🟩",
    "DUNAS": "🟨",
    "DUNAS2": "🟪",
    "FOZ": "🟧",
    "ESCAPE": "🟫",
    "MARES": "⬛",
}

ALOJAMENTO_BUTTON_COLORS = {
    "ABH": "#ef4444",
    "AFH": "#3b82f6",
    "PIPO": "#22c55e",
    "DUNAS": "#facc15",
    "DUNAS2": "#a855f7",
    "FOZ": "#fb923c",
    "ESCAPE": "#92400e",
    "MARES": "#374151",
}


def format_alojamento_badge(alojamento_value):
    if pd.isna(alojamento_value):
        return "⬜"
    alojamento_txt = str(alojamento_value).strip()
    badge = ALOJAMENTO_BADGES.get(alojamento_txt.upper(), "⬜")
    return f"{badge} {alojamento_txt}" if alojamento_txt else "⬜"


def extract_room_tag(unidade_value):
    import re

    if pd.isna(unidade_value):
        return ""

    raw = str(unidade_value).strip()
    if not raw:
        return ""

    lowered = raw.lower()
    patterns = [
        (r"\bquarto\b[^\d]{0,30}n?[ºo]?\s*(\d+)\b", "Q"),
        (r"\bq\s*[-:#]?\s*(\d+)\b", "Q"),
        (r"\bcama\b[^\d]{0,30}n?[ºo]?\s*(\d+)\b", "C"),
        (r"\bc\s*[-:#]?\s*(\d+)\b", "C"),
        (r"\bapartamento\b[^\d]{0,30}n?[ºo]?\s*([a-z0-9]+)\b", "A"),
        (r"\b(?:apto|apt\.?)\s*[-:#]?\s*([a-z0-9]+)\b", "A"),
    ]

    for pattern, prefix in patterns:
        match = re.search(pattern, lowered, flags=re.IGNORECASE)
        if match:
            return f"{prefix}{match.group(1).upper()}"

    unidades = parse_unidade_labels(unidade_value)
    if not unidades:
        return ""

    primeira = unidades[0]
    if primeira.lower().startswith("quarto "):
        room_match = re.search(r"(\d+)", primeira)
        return f"Q{room_match.group(1)}" if room_match else ""
    if primeira.lower().startswith("cama "):
        bed_match = re.search(r"(\d+)", primeira)
        return f"C{bed_match.group(1)}" if bed_match else ""
    if primeira.lower().startswith("apartamento "):
        apt_match = re.search(r"([a-z0-9]+)", primeira.split(" ", 1)[1], flags=re.IGNORECASE)
        return f"A{apt_match.group(1).upper()}" if apt_match else ""
    return ""


def format_nome_com_quarto(nome_value, unidade_value):
    nome = "" if pd.isna(nome_value) else str(nome_value).strip()
    # Usar exatamente o mesmo formato da coluna 'Unidade' das reservas
    quartos_str = format_quartos_text(unidade_value)
    if nome and quartos_str and quartos_str != "Sem unidade definida":
        return f"{nome} ({quartos_str})"
    if quartos_str and quartos_str != "Sem unidade definida":
        return quartos_str
    return nome


def parse_unidade_labels(unidade_value):
    import re

    if pd.isna(unidade_value):
        return []

    raw = str(unidade_value).strip()
    if not raw:
        return []

    partes = [p.strip() for p in re.split(r"[,;|]", raw) if p.strip()]
    labels = []
    seen = set()

    def _add(label):
        if label and label not in seen:
            seen.add(label)
            labels.append(label)

    def _split_ids(ids_chunk):
        return [t.strip() for t in re.split(r"\s*(?:e|/|-)\s*", ids_chunk) if t.strip()]

    for parte in partes:
        p = parte.lower().strip()

        # Ignora valores genéricos que não identificam unidade real.
        if re.fullmatch(r"\d+(?:\.0+)?", p):
            continue
        if re.fullmatch(r"unidade\s*[:#-]?\s*\d+(?:\.0+)?", p):
            continue

        # Ignora contagens genéricas sem unidade concreta (ex.: "2 quartos").
        if re.fullmatch(r"\d+\s+(quartos?|camas?|apartamentos?)", p):
            continue

        found = False

        # Captura formatos como "Quarto Duplo nº1" para não perder a unidade.
        explicit_patterns = [
            (r"\bquartos?\b.{0,40}?n?[ºo]?\s*(\d+)\b", "Quarto"),
            (r"\bcamas?\b.{0,40}?n?[ºo]?\s*(\d+)\b", "Cama"),
            (r"\bapartamentos?\b.{0,40}?n?[ºo]?\s*([a-z0-9]+)\b", "Apartamento"),
            (r"\b(?:apto|apt\.?)\b.{0,40}?n?[ºo]?\s*([a-z0-9]+)\b", "Apartamento"),
        ]
        for pattern, prefix in explicit_patterns:
            for m in re.finditer(pattern, p, flags=re.IGNORECASE):
                _add(f"{prefix} {m.group(1).upper()}")
                found = True

        patterns = [
            (r"apartamentos?\s*([a-z0-9]+(?:\s*(?:e|/|-)\s*[a-z0-9]+)*)", "Apartamento"),
            (r"(?:apto|apt\.?)\s*([a-z0-9]+(?:\s*(?:e|/|-)\s*[a-z0-9]+)*)", "Apartamento"),
            (r"quartos?\s*([a-z0-9]+(?:\s*(?:e|/|-)\s*[a-z0-9]+)*)", "Quarto"),
            (r"camas?\s*([a-z0-9]+(?:\s*(?:e|/|-)\s*[a-z0-9]+)*)", "Cama"),
        ]

        for pattern, prefix in patterns:
            for m in re.finditer(pattern, p, flags=re.IGNORECASE):
                for unit_id in _split_ids(m.group(1)):
                    _add(f"{prefix} {unit_id.upper()}")
                    found = True

        if not found:
            # Só mantém fallback quando há texto útil; evita mostrar apenas números.
            if re.fullmatch(r"\d+(?:\.0+)?", p):
                continue
            _add(parte)

    return labels


def format_quartos_text(unidade_value):
    import re

    labels = parse_unidade_labels(unidade_value)
    if not labels:
        return "Sem unidade definida"

    short_labels = []
    seen = set()

    def _push(tag):
        if tag and tag not in seen:
            seen.add(tag)
            short_labels.append(tag)

    for label in labels:
        txt = str(label).strip()
        if not txt:
            continue

        matches = [
            (r"\bquarto\b.{0,40}?n?[ºo]?\s*(\d+)\b", "Q"),
            (r"\bq\s*[-:#]?\s*(\d+)\b", "Q"),
            (r"\bcama\b.{0,40}?n?[ºo]?\s*(\d+)\b", "C"),
            (r"\bc\s*[-:#]?\s*(\d+)\b", "C"),
            (r"\bapartamento\b.{0,40}?n?[ºo]?\s*([a-z0-9]+)\b", "A"),
            (r"\b(?:apto|apt\.?)\s*[-:#]?\s*([a-z0-9]+)\b", "A"),
        ]

        normalized_tag = ""
        for pattern, prefix in matches:
            m = re.search(pattern, txt, flags=re.IGNORECASE)
            if m:
                normalized_tag = f"{prefix}{m.group(1).upper()}"
                break

        _push(normalized_tag)

    if short_labels:
        return ", ".join(short_labels)

    # Fallback: mostra as unidades originais quando não foi possível abreviar.
    return ", ".join(labels) if labels else "Sem unidade definida"


def _build_quick_access_button_css(alojamentos):
    css_parts = []
    for aloj in alojamentos:
        color = ALOJAMENTO_BUTTON_COLORS.get(str(aloj).upper(), "#9ca3af")
        css_parts.append(
            f"""
            div[data-testid=\"stVerticalBlock\"] div[data-testid=\"stButton\"] button[kind=\"secondary\"][id$=\"quick_aloj_{str(aloj).lower()}\"] {{
                background-color: {color} !important;
                color: #ffffff !important;
                border: 1px solid {color} !important;
                font-weight: 700 !important;
            }}
            """
        )

    if css_parts:
        st.markdown(f"<style>{''.join(css_parts)}</style>", unsafe_allow_html=True)


def render_quick_access_tab(df, suggested_times):
    if df.empty:
        st.info("Sem dados ainda. Importa ficheiros no separador 'Importar' para começar.")
        return

    quick_df = sanitize_optional_columns(df.copy())
    quick_df["_aloj_norm"] = quick_df["Alojamento"].apply(normalize_alojamento)
    alojamentos = (
        quick_df["_aloj_norm"]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    )
    alojamentos = sorted(alojamentos)

    if not alojamentos:
        st.info("Sem alojamentos disponíveis nas reservas atuais.")
        return

    if "quick_selected_aloj" not in st.session_state:
        st.session_state["quick_selected_aloj"] = None
    if "quick_selected_idx" not in st.session_state:
        st.session_state["quick_selected_idx"] = None

    if st.session_state["quick_selected_aloj"] not in alojamentos:
        st.session_state["quick_selected_aloj"] = None
        st.session_state["quick_selected_idx"] = None

    st.subheader("Acesso rápido para check-ins")

    quartos_hoje = st.session_state.get("quartos_disponiveis", [])
    if "quick_inserir_quarto_idx" not in st.session_state:
        st.session_state["quick_inserir_quarto_idx"] = None

    _build_quick_access_button_css(alojamentos)

    selected_aloj = st.session_state["quick_selected_aloj"]
    selected_idx = st.session_state["quick_selected_idx"]

    # Passo 1: apenas alojamentos
    if not selected_aloj:
        cols_aloj = st.columns(min(4, len(alojamentos)))
        for i, aloj in enumerate(alojamentos):
            badge = ALOJAMENTO_BADGES.get(aloj, "⬜")
            with cols_aloj[i % len(cols_aloj)]:
                if st.button(
                    f"{badge} {aloj}",
                    key=f"quick_aloj_{aloj.lower()}",
                    use_container_width=True,
                ):
                    st.session_state["quick_selected_aloj"] = aloj
                    st.session_state["quick_selected_idx"] = None
                    st.rerun()

        if quartos_hoje:
            st.divider()
            st.markdown("**Quartos disponíveis hoje**")
            _today = datetime.now().date()
            _tomorrow = _today + timedelta(days=1)
            _hora_opts = ["nenhuma"] + build_suggested_times()
            for _qi, _q in enumerate(quartos_hoje):
                _badge = ALOJAMENTO_BADGES.get(_q["alojamento"].upper(), "⬜")
                _preco = _q.get("preco", 0)
                _preco_str = f"{int(_preco)} €" if _preco == int(_preco) else f"{_preco:.2f} €"
                _pax_str = f" · {_q['pessoas']} pax" if _q.get("pessoas") else ""
                _nota_str = f" · {_q['notas']}" if _q.get("notas") else ""
                _is_open = st.session_state["quick_inserir_quarto_idx"] == _qi
                _row_label = f"{_badge} {_q['alojamento']} — {_q['unidade']} — {_preco_str}{_pax_str}{_nota_str}"
                if st.button(
                    f"▾ {_row_label}" if _is_open else _row_label,
                    key=f"qins_{_qi}",
                    use_container_width=True,
                ):
                    if _is_open:
                        st.session_state["quick_inserir_quarto_idx"] = None
                    else:
                        st.session_state["quick_inserir_quarto_idx"] = _qi
                    st.rerun()
                if st.session_state["quick_inserir_quarto_idx"] == _qi:
                    with st.form(key=f"qins_form_{_qi}", clear_on_submit=True):
                        _fi1, _fi2 = st.columns(2)
                        with _fi1:
                            _ins_nome = st.text_input("Nome do hóspede")
                        with _fi2:
                            _ins_pessoas = st.number_input("Pessoas", min_value=1, step=1, value=int(_q.get("pessoas") or 1))
                        _fi3, _fi4 = st.columns(2)
                        with _fi3:
                            _ins_hora = st.selectbox("Hora PA", options=_hora_opts, index=0)
                        with _fi4:
                            _ins_notas = st.text_input("Notas", placeholder="Escreve uma nota...")
                        _ins_submit = st.form_submit_button("Guardar reserva")
                    if _ins_submit:
                        _ins_nome_txt = str(_ins_nome).strip()
                        if not _ins_nome_txt:
                            show_pink_alert("Indica o nome do hóspede.")
                        else:
                            _base = st.session_state.get("reservas_df", pd.DataFrame()).copy()
                            _nova_reserva = {
                                "Nome": _ins_nome_txt,
                                "Alojamento": _q["alojamento"],
                                "Unidade": _q["unidade"],
                                "Pessoas": int(_ins_pessoas),
                                "Check-in": _today,
                                "Check-out": _tomorrow,
                                "Origem": "Direta",
                                "Hora PA": None if _ins_hora == "nenhuma" else _ins_hora,
                                "Notas": str(_ins_notas).strip() or None,
                            }
                            if _is_duplicate_reserva(_base, _nova_reserva):
                                show_pink_alert(
                                    "Já existe uma reserva igual para este hóspede. Verifica nome/datas/unidade antes de guardar."
                                )
                            else:
                                _nova = pd.DataFrame([_nova_reserva])
                                _merged, _ = merge_new_reservas(_base, _nova)
                                _merged = sanitize_optional_columns(_merged)
                                st.session_state["reservas_df"] = _merged
                                st.session_state["reservas_editor_df"] = _merged
                                save_reservas(_merged)
                                st.session_state["saidas_checklist"] = load_saidas_checklist(_merged)
                                st.session_state["quartos_disponiveis"].pop(_qi)
                                save_quartos(st.session_state["quartos_disponiveis"])
                                st.session_state["quick_inserir_quarto_idx"] = None
                                st.success(f"{_ins_nome_txt} inserido no {_q['unidade']} ({_q['alojamento']}).")
                                st.rerun()
        return

    pessoas_df = quick_df[quick_df["_aloj_norm"] == selected_aloj].copy()
    pessoas_df = pessoas_df.reset_index().rename(columns={"index": "_idx"})

    # Passo 2: apenas nomes do alojamento selecionado
    if selected_idx is None:
        ctop1, ctop2 = st.columns([1.5, 4])
        with ctop1:
            if st.button("← Voltar aos alojamentos", key="quick_back_to_aloj"):
                st.session_state["quick_selected_aloj"] = None
                st.session_state["quick_selected_idx"] = None
                st.rerun()
        with ctop2:
            safe_aloj = html.escape(str(selected_aloj))
            safe_badge = html.escape(ALOJAMENTO_BADGES.get(str(selected_aloj).upper(), "⬜"))
            st.markdown(
                f"<div style='margin-top:0.45rem;font-weight:700;'>Alojamento selecionado: {safe_badge} {safe_aloj}</div>",
                unsafe_allow_html=True,
            )

        if pessoas_df.empty:
            st.info("Sem pessoas neste alojamento.")
            return

        st.caption("Seleciona a pessoa")
        cols_pessoas = st.columns(3)
        for i, (_, row) in enumerate(pessoas_df.iterrows()):
            nome_botao = format_nome_com_quarto(row.get("Nome"), row.get("Unidade")) or "Sem nome"
            with cols_pessoas[i % 3]:
                if st.button(
                    nome_botao,
                    key=f"quick_guest_{selected_aloj.lower()}_{int(row['_idx'])}",
                    use_container_width=True,
                ):
                    st.session_state["quick_selected_idx"] = int(row["_idx"])
                    st.rerun()
        return

    if selected_idx not in quick_df.index:
        st.session_state["quick_selected_idx"] = None
        st.rerun()

    row = quick_df.loc[selected_idx]

    # Passo 3: apenas ficha da pessoa selecionada
    ctop1, ctop2 = st.columns([1.5, 4])
    with ctop1:
        if st.button("← Voltar aos nomes", key="quick_back_to_names"):
            st.session_state["quick_selected_idx"] = None
            st.rerun()
    with ctop2:
        safe_aloj = html.escape(str(selected_aloj))
        safe_badge = html.escape(ALOJAMENTO_BADGES.get(str(selected_aloj).upper(), "⬜"))
        st.markdown(
            f"<div style='margin-top:0.45rem;font-weight:700;'>Alojamento: {safe_badge} {safe_aloj}</div>",
            unsafe_allow_html=True,
        )

    nome_sel = "Sem nome" if pd.isna(row.get("Nome")) else str(row.get("Nome")).strip()
    quartos_text = format_quartos_text(row.get("Unidade"))
    st.markdown(f"### {nome_sel}")
    st.success(f"Unidade(s): {quartos_text}")

    hora_options = ["nenhuma"] + suggested_times
    pago_options = ["Não", "Sim"]

    hora_default = str(row.get("Hora PA")) if pd.notna(row.get("Hora PA")) else None
    if hora_default not in hora_options:
        hora_default = None

    pago_default = str(row.get("PA pago")) if pd.notna(row.get("PA pago")) else None
    if pago_default not in pago_options:
        pago_default = None

    notas_default = str(row.get("Notas")) if pd.notna(row.get("Notas")) else ""

    c1, c2, c3 = st.columns([1, 1, 1.5])
    with c1:
        nova_hora = st.selectbox(
            "Hora PA",
            options=hora_options,
            index=hora_options.index(hora_default) if hora_default is not None else None,
            placeholder="nenhuma",
            key=f"quick_hora_pa_{selected_idx}",
        )
    with c2:
        novo_pago = st.selectbox(
            "PA pago",
            options=pago_options,
            index=pago_options.index(pago_default) if pago_default is not None else None,
            placeholder="Não",
            key=f"quick_pa_pago_{selected_idx}",
        )
    with c3:
        novas_notas = st.text_input(
            "Notas",
            value=notas_default,
            key=f"quick_notas_{selected_idx}",
            placeholder="Escreve uma nota...",
        )

    if st.button("Salvar alterações", key=f"quick_save_{selected_idx}", type="primary"):
        updated_df = quick_df.copy()
        updated_df.loc[selected_idx, "Hora PA"] = None if not nova_hora or nova_hora == "nenhuma" else nova_hora
        updated_df.loc[selected_idx, "PA pago"] = None if not novo_pago or novo_pago == "Não" else novo_pago
        updated_df.loc[selected_idx, "Notas"] = None if not str(novas_notas).strip() else str(novas_notas).strip()
        updated_df = sanitize_optional_columns(updated_df)

        st.session_state["reservas_editor_df"] = updated_df.copy()
        st.session_state["reservas_df"] = updated_df.copy()
        save_reservas(updated_df)
        st.success("Alterações guardadas com sucesso.")
        st.rerun()

def _render_reservas_editor_impl(suggested_times):
    editor_df = sanitize_optional_columns(
        st.session_state.get("reservas_editor_df", pd.DataFrame()).copy()
    )

    if editor_df.empty:
        st.info("Sem dados ainda. Importa ficheiros no separador 'Importar' para começar.")
        return

    conflicts = detect_conflicts(editor_df)
    if conflicts:
        with st.expander(f"⚠ {len(conflicts)} conflito(s) de reservas detectado(s)", expanded=True):
            for c in conflicts:
                st.markdown(c)

    vista_compacta = st.toggle("Vista compacta", value=False, key="reservas_vista_compacta")

    display_df = editor_df.copy()
    if "Alojamento" in display_df.columns:
        display_df["Alojamento"] = display_df["Alojamento"].apply(format_alojamento_badge)
    if "Nome" in display_df.columns:
        display_df["Nome"] = display_df["Nome"].fillna("")
    if "Unidade" in display_df.columns:
        display_df["Unidade"] = display_df["Unidade"].apply(
            lambda v: "" if pd.isna(v) or not str(v).strip() else format_quartos_text(v)
        )
    display_df["Check-in/Check-out"] = display_df.apply(
        lambda row: format_checkin_checkout(row.get("Check-in"), row.get("Check-out")),
        axis=1,
    )
    display_df = display_df.drop(columns=[c for c in ["Check-in", "Check-out"] if c in display_df.columns])
    if vista_compacta:
        ordered_display_cols = ["Alojamento", "Nome", "Unidade", "Pessoas", "Check-in/Check-out"]
    else:
        ordered_display_cols = [
            "Alojamento",
            "Nome",
            "Unidade",
            "Pessoas",
            "Check-in/Check-out",
            "Hora PA",
            "PA pago",
            "Notas",
        ]
    display_df = display_df[[c for c in ordered_display_cols if c in display_df.columns]]

    _fill_cols = [c for c in display_df.columns if c not in ("Hora PA", "PA pago", "Notas")]
    display_df[_fill_cols] = display_df[_fill_cols].fillna("")
    # Normaliza unidades "casa inteira" para leitura humana na vista
    _CASA_INTEIRA_RE = re.compile(
        r'\b(two[\s-]bedroom|three[\s-]bedroom|four[\s-]bedroom'
        r'|house|entire|whole|completo|inteiro'
        r'|casa inteira|apartamento inteiro)\b',
        re.IGNORECASE,
    )
    if "Unidade" in display_df.columns:
        display_df["Unidade"] = display_df["Unidade"].apply(
            lambda v: "Apartamento inteiro" if isinstance(v, str) and _CASA_INTEIRA_RE.search(v) else v
        )
    # Sem scroll interno: usa apenas o scroll da página.
    editor_height = max(240, 42 + (len(display_df) + 1) * 35)
    edited_from_table = st.data_editor(
        display_df,
        key="reservas_editor",
        width='stretch',
        height=editor_height,
        hide_index=True,
        disabled=["Nome", "Check-in/Check-out", "Pessoas", "Unidade", "Alojamento"],
        column_config={
            "Hora PA": st.column_config.SelectboxColumn(
                "Hora pequeno-almoço",
                options=suggested_times,
                required=False,
                help="Se ficar vazio/cinzento, significa nenhuma hora definida.",
            ),
            "PA pago": st.column_config.SelectboxColumn(
                "PA PAGO",
                options=["Sim"],
                required=False,
                help="Se ficar vazio/cinzento, significa Não.",
            ),
            "Notas": st.column_config.TextColumn(
                "NOTAS",
                help="Se ficar vazio/cinzento, significa sem nota.",
                width="medium",
            ),
        }
    )

    editable_cols = [c for c in ["Hora PA", "PA pago", "Notas"] if c in edited_from_table.columns]
    before_str = editor_df[editable_cols].astype(str).values

    updated_df = editor_df.copy()
    for col in editable_cols:
        updated_df[col] = edited_from_table[col]
    updated_df = sanitize_optional_columns(updated_df)

    after_str = updated_df[editable_cols].astype(str).values
    if not (before_str == after_str).all():
        old_overcrowding = set(build_overcrowding_messages(editor_df, suggested_times))
        new_overcrowding = set(build_overcrowding_messages(updated_df, suggested_times))
        created_overcrowding = sorted(new_overcrowding - old_overcrowding)
        if created_overcrowding:
            st.session_state["pending_overcrowding_messages"] = created_overcrowding
            st.session_state["show_overcrowding_ack"] = True

        st.session_state["reservas_editor_df"] = updated_df.copy()
        st.session_state["reservas_df"] = updated_df.copy()
        save_reservas(updated_df)
        st.session_state["saidas_checklist"] = load_saidas_checklist(updated_df)
        editor_df = updated_df

    st.divider()

    select_df = editor_df.reset_index().rename(columns={"index": "_idx"})
    reserva_idx = st.selectbox(
        "Selecionar reserva",
        options=select_df["_idx"].tolist(),
        format_func=lambda i: (
            f"{'' if pd.isna(select_df.loc[select_df['_idx'] == i, 'Nome'].iloc[0]) else select_df.loc[select_df['_idx'] == i, 'Nome'].iloc[0]} | "
            f"{'' if pd.isna(select_df.loc[select_df['_idx'] == i, 'Alojamento'].iloc[0]) else select_df.loc[select_df['_idx'] == i, 'Alojamento'].iloc[0]} "
            f"{'' if pd.isna(select_df.loc[select_df['_idx'] == i, 'Unidade'].iloc[0]) else format_quartos_text(select_df.loc[select_df['_idx'] == i, 'Unidade'].iloc[0])}"
        ),
        key="reserva_idx_select",
    )

    current_hora = editor_df.loc[reserva_idx, "Hora PA"]
    current_pago = editor_df.loc[reserva_idx, "PA pago"]
    current_notas = editor_df.loc[reserva_idx, "Notas"] if "Notas" in editor_df.columns else None

    hora_options = ["nenhuma"] + suggested_times
    pago_options = ["Não", "Sim"]

    hora_default = str(current_hora) if pd.notna(current_hora) else None
    if hora_default not in hora_options:
        hora_default = None

    pago_default = str(current_pago) if pd.notna(current_pago) else None
    if pago_default not in pago_options:
        pago_default = None

    notas_default = str(current_notas) if pd.notna(current_notas) else ""

    unidade_default = "" if pd.isna(editor_df.loc[reserva_idx, "Unidade"]) else str(editor_df.loc[reserva_idx, "Unidade"]).strip()
    pessoas_default_num = pd.to_numeric(editor_df.loc[reserva_idx, "Pessoas"], errors="coerce")
    pessoas_default = int(pessoas_default_num) if pd.notna(pessoas_default_num) and int(pessoas_default_num) > 0 else 1

    # Recarrega os campos sempre que muda a reserva selecionada para evitar estado "preso".
    if st.session_state.get("reserva_form_loaded_idx") != reserva_idx:
        st.session_state["reserva_form_loaded_idx"] = reserva_idx
        st.session_state["reserva_form_hora"] = hora_default
        st.session_state["reserva_form_pago"] = pago_default
        st.session_state["reserva_form_pessoas"] = pessoas_default
        st.session_state["reserva_form_notas"] = notas_default
        st.session_state["reserva_form_unidade"] = unidade_default

    c1, c2, c3, c4, c5 = st.columns([1, 1, 1.2, 1.5, 1.4])
    with c1:
        st.selectbox(
            "Hora PA",
            options=hora_options,
            index=hora_options.index(st.session_state.get("reserva_form_hora"))
            if st.session_state.get("reserva_form_hora") in hora_options
            else None,
            placeholder="nenhuma",
            key="reserva_form_hora",
        )
    with c2:
        st.selectbox(
            "PA pago",
            options=pago_options,
            index=pago_options.index(st.session_state.get("reserva_form_pago"))
            if st.session_state.get("reserva_form_pago") in pago_options
            else None,
            placeholder="Não",
            key="reserva_form_pago",
        )
    with c3:
        st.number_input(
            "Pessoas",
            min_value=1,
            step=1,
            value=int(st.session_state.get("reserva_form_pessoas", pessoas_default)),
            key="reserva_form_pessoas",
        )
    with c4:
        st.text_input(
            "Notas",
            value=st.session_state.get("reserva_form_notas", notas_default),
            key="reserva_form_notas",
            placeholder="Escreve uma nota...",
        )
    with c5:
        st.text_input(
            "Quarto/Unidade",
            value=st.session_state.get("reserva_form_unidade", unidade_default),
            key="reserva_form_unidade",
            placeholder="Ex: Quarto 2",
        )

    if st.button("Aplicar alterações à reserva", key="apply_reserva_update"):
        df_before_update = editor_df.copy()
        nova_hora = st.session_state.get("reserva_form_hora")
        novo_pago = st.session_state.get("reserva_form_pago")
        novas_pessoas = st.session_state.get("reserva_form_pessoas", pessoas_default)
        novas_notas = st.session_state.get("reserva_form_notas", "")
        nova_unidade = st.session_state.get("reserva_form_unidade", "")
        editor_df.loc[reserva_idx, "Hora PA"] = None if not nova_hora or nova_hora == "nenhuma" else nova_hora
        editor_df.loc[reserva_idx, "PA pago"] = None if not novo_pago or novo_pago == "Não" else novo_pago
        editor_df.loc[reserva_idx, "Pessoas"] = int(novas_pessoas)
        editor_df.loc[reserva_idx, "Notas"] = None if not str(novas_notas).strip() or str(novas_notas).strip().lower() == "nenhuma" else novas_notas
        editor_df.loc[reserva_idx, "Unidade"] = None if not str(nova_unidade).strip() else str(nova_unidade).strip()
        editor_df = sanitize_optional_columns(editor_df)

        old_overcrowding = set(build_overcrowding_messages(df_before_update, suggested_times))
        new_overcrowding = set(build_overcrowding_messages(editor_df, suggested_times))
        created_overcrowding = sorted(new_overcrowding - old_overcrowding)
        if created_overcrowding:
            st.session_state["pending_overcrowding_messages"] = created_overcrowding
            st.session_state["show_overcrowding_ack"] = True

        st.session_state["reservas_editor_df"] = editor_df.copy()
        st.session_state["reservas_df"] = editor_df.copy()
        save_reservas(editor_df)
        st.session_state["saidas_checklist"] = load_saidas_checklist(editor_df)
        st.success("Reserva atualizada.")
        st.rerun()

    if "delete_reserva_confirm_idx" not in st.session_state:
        st.session_state["delete_reserva_confirm_idx"] = None

    d1, d2 = st.columns([1.2, 4])
    with d1:
        if st.button("Eliminar reserva", key="delete_reserva_selected", type="secondary"):
            st.session_state["delete_reserva_confirm_idx"] = reserva_idx

    if st.session_state.get("delete_reserva_confirm_idx") == reserva_idx:
        show_pink_alert("Confirma eliminação da reserva selecionada?")
        dc1, dc2 = st.columns([1.2, 1.2])
        with dc1:
            if st.button("Confirmar eliminação", key=f"confirm_delete_{reserva_idx}", type="primary"):
                updated_df = editor_df.drop(index=reserva_idx).reset_index(drop=True)
                updated_df = sanitize_optional_columns(updated_df)
                st.session_state["reservas_editor_df"] = updated_df.copy()
                st.session_state["reservas_df"] = updated_df.copy()
                save_reservas(updated_df)
                st.session_state["saidas_checklist"] = load_saidas_checklist(updated_df)
                st.session_state["delete_reserva_confirm_idx"] = None
                st.success("Reserva eliminada com sucesso.")
                st.rerun()
        with dc2:
            if st.button("Cancelar eliminação", key=f"cancel_delete_{reserva_idx}"):
                st.session_state["delete_reserva_confirm_idx"] = None

    if st.checkbox("Editar todos os campos da reserva selecionada", value=False, key="show_full_edit"):
        row_edit = editor_df.loc[reserva_idx]
        checkin_edit_options = _existing_date_options(editor_df, "Check-in")
        checkout_edit_options = _existing_date_options(editor_df, "Check-out")

        with st.form("editar_reserva_form"):
            ef1, ef2, ef3 = st.columns(3)
            with ef1:
                edit_nome = st.text_input("Nome", value=str(row_edit.get("Nome", "") or ""))
            with ef2:
                aloj_atual = str(row_edit.get("Alojamento", "") or "")
                aloj_idx = ALOJAMENTOS.index(aloj_atual) if aloj_atual in ALOJAMENTOS else 0
                edit_alojamento = st.selectbox("Alojamento", ALOJAMENTOS, index=aloj_idx)
            with ef3:
                edit_pessoas = st.number_input(
                    "Pessoas",
                    min_value=1,
                    step=1,
                    value=int(pd.to_numeric(row_edit.get("Pessoas", 1), errors="coerce") or 1),
                )

            ef4, ef5, ef6 = st.columns(3)
            with ef4:
                checkin_atual = pd.to_datetime(row_edit.get("Check-in"), errors="coerce")
                checkin_atual_date = checkin_atual.date() if pd.notna(checkin_atual) else datetime.now().date()
                edit_checkin = _choose_date_with_suggestions(
                    "Check-in",
                    checkin_edit_options,
                    checkin_atual_date,
                    key_prefix=f"edit_reserva_checkin_{reserva_idx}",
                )
            with ef5:
                checkout_atual = pd.to_datetime(row_edit.get("Check-out"), errors="coerce")
                checkout_atual_date = checkout_atual.date() if pd.notna(checkout_atual) else checkin_atual_date
                edit_checkout = _choose_date_with_suggestions(
                    "Check-out",
                    checkout_edit_options,
                    checkout_atual_date,
                    key_prefix=f"edit_reserva_checkout_{reserva_idx}",
                )
            with ef6:
                edit_unidade = st.text_input("Unidade", value=str(row_edit.get("Unidade", "") or ""))

            ef7, ef8, ef9 = st.columns(3)
            hora_opts_edit = ["nenhuma"] + suggested_times
            hora_atual_edit = str(row_edit.get("Hora PA")) if pd.notna(row_edit.get("Hora PA")) else "nenhuma"
            if hora_atual_edit not in hora_opts_edit:
                hora_atual_edit = "nenhuma"
            pago_opts_edit = ["Não", "Sim"]
            pago_atual_edit = str(row_edit.get("PA pago")) if pd.notna(row_edit.get("PA pago")) else "Não"
            if pago_atual_edit not in pago_opts_edit:
                pago_atual_edit = "Não"
            with ef7:
                edit_hora_pa = st.selectbox(
                    "Hora PA",
                    options=hora_opts_edit,
                    index=hora_opts_edit.index(hora_atual_edit),
                )
            with ef8:
                edit_pa_pago = st.selectbox(
                    "PA pago",
                    options=pago_opts_edit,
                    index=pago_opts_edit.index(pago_atual_edit),
                )
            with ef9:
                edit_notas = st.text_input(
                    "Notas",
                    value=str(row_edit.get("Notas", "") or ""),
                    placeholder="Escreve uma nota...",
                )

            edit_submit = st.form_submit_button("Guardar alterações completas")

        if edit_submit:
            nome_edit_txt = str(edit_nome).strip()
            if not nome_edit_txt:
                show_pink_alert("O nome não pode ficar vazio.")
            elif edit_checkout < edit_checkin:
                show_pink_alert("A data de check-out não pode ser anterior ao check-in.")
            else:
                checkin_store = pd.to_datetime(edit_checkin, errors="coerce")
                checkout_store = pd.to_datetime(edit_checkout, errors="coerce")
                checkin_store = checkin_store.strftime("%Y-%m-%d") if pd.notna(checkin_store) else None
                checkout_store = checkout_store.strftime("%Y-%m-%d") if pd.notna(checkout_store) else None

                _edit_candidate = {
                    "Nome": nome_edit_txt,
                    "Alojamento": edit_alojamento,
                    "Unidade": str(edit_unidade).strip() or None,
                    "Pessoas": int(edit_pessoas),
                    "Check-in": checkin_store,
                    "Check-out": checkout_store,
                }
                if _is_duplicate_reserva(editor_df, _edit_candidate, ignore_index=reserva_idx):
                    show_pink_alert(
                        "Esta edição criaria uma reserva duplicada (mesmo nome, alojamento, unidade, datas e número de pessoas já existe). Verifica os dados antes de guardar."
                    )
                else:
                    editor_df.loc[reserva_idx, "Nome"] = nome_edit_txt
                    editor_df.loc[reserva_idx, "Alojamento"] = edit_alojamento
                    editor_df.loc[reserva_idx, "Pessoas"] = int(edit_pessoas)
                    editor_df.loc[reserva_idx, "Check-in"] = checkin_store
                    editor_df.loc[reserva_idx, "Check-out"] = checkout_store
                    editor_df.loc[reserva_idx, "Unidade"] = str(edit_unidade).strip() or None
                    editor_df.loc[reserva_idx, "Hora PA"] = None if edit_hora_pa == "nenhuma" else edit_hora_pa
                    editor_df.loc[reserva_idx, "PA pago"] = "Sim" if edit_pa_pago == "Sim" else None
                    editor_df.loc[reserva_idx, "Notas"] = str(edit_notas).strip() or None
                    editor_df = sanitize_optional_columns(editor_df)
                    st.session_state["reservas_editor_df"] = editor_df.copy()
                    st.session_state["reservas_df"] = editor_df.copy()
                    save_reservas(editor_df)
                    st.session_state["saidas_checklist"] = load_saidas_checklist(editor_df)
                    st.success("Reserva editada com sucesso.")
                    st.rerun()

    st.divider()
    st.subheader(" Ocupação desta noite")
    st.metric("Total de hóspedes", total_hospedes_esta_noite(editor_df))

    st.divider()
    st.caption("Exportação")

    from io import BytesIO

    excel_buffer = BytesIO()
    display_df_export = sanitize_optional_columns(editor_df).fillna("")
    display_df_export.to_excel(excel_buffer, index=False, engine="openpyxl", sheet_name="Reservas")
    excel_buffer.seek(0)
    st.download_button(
        label=" Exportar para Excel",
        data=excel_buffer.getvalue(),
        file_name="reservas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


render_reservas_editor = (
    st.fragment(_render_reservas_editor_impl)
    if hasattr(st, "fragment")
    else _render_reservas_editor_impl
)


if "reservas_df" not in st.session_state:
    st.session_state["reservas_df"] = load_reservas()
if "pending_overcrowding_messages" not in st.session_state:
    st.session_state["pending_overcrowding_messages"] = []
if "show_overcrowding_ack" not in st.session_state:
    st.session_state["show_overcrowding_ack"] = False
if "quartos_disponiveis" not in st.session_state:
    st.session_state["quartos_disponiveis"] = load_quartos()
if "notas_gerais_pa" not in st.session_state:
    st.session_state["notas_gerais_pa"] = load_notas_gerais()
if "notas_gerais_pa_editor" not in st.session_state:
    st.session_state["notas_gerais_pa_editor"] = st.session_state["notas_gerais_pa"]
if "transfers" not in st.session_state:
    st.session_state["transfers"] = load_transfers()

header_c1, header_c2 = st.columns([4, 1.2])
with header_c1:
    st.markdown("### Gestão de Reservas + Pequenos-Almoços")
with header_c2:
    if st.button(" Atualizar", use_container_width=True, help="Recarrega dados guardados do ficheiro/Supabase"):
        refresh_data_from_storage()
        st.success("Dados atualizados.")
        st.rerun()

st.caption(f"Guardado pela última vez: {get_last_saved_text()}")

if st.session_state["show_overcrowding_ack"] and st.session_state["pending_overcrowding_messages"]:
    show_pink_alert("Foram criados horários com sobrelotação. Confirma para continuar.")
    for msg in st.session_state["pending_overcrowding_messages"]:
        show_pink_alert(msg)
    if st.button("OK", key="ack_overcrowding"):
        st.session_state["show_overcrowding_ack"] = False
        st.session_state["pending_overcrowding_messages"] = []
        st.rerun()

tab_acesso_rapido, tab_reservas, tab_pa, tab_saidas, tab_notas, tab_inserir, tab_importar = st.tabs(
    ["Acesso rápido", "Reservas", "Pequenos almoços", "Saídas", "Notas", "Inserir", "Importar"]
)

if "reservas_df" not in st.session_state or not isinstance(st.session_state["reservas_df"], pd.DataFrame) or st.session_state["reservas_df"].empty:
    st.warning("⚠️ Ainda não existem dados inseridos ou importados. Usa o separador 'Importar' ou 'Inserir' para começar.")

# --- Checklist de Saídas ---
with tab_saidas:
    st.header("Checklist de Saídas (Limpezas)")
    st.info("Visualize e confirme as saídas previstas para amanhã. Marque/desmarque manualmente conforme necessário.")

    # --- Estrutura fixa dos alojamentos e quartos ---
    CHECKLIST_STRUCTURE = [
        ("ABH", [
            "Quarto 1", "Quarto 2", "Quarto 3", "Quarto 4 - Cama 1", "Quarto 4 - Cama 2", "Quarto 4 - Cama 3", "Quarto 4 - Cama 4", "Quarto 5"
        ]),
        ("AFH", [
            "Quarto 1", "Quarto 2", "Quarto 3", "Quarto 4", "Quarto 5", "Quarto 6"
        ]),
        ("DUNAS", [
            "Quarto 1", "Quarto 2", "Quarto 3"
        ]),
        ("PIPO", [
            "Quarto 1", "Quarto 2"
        ]),
        ("ESCAPE", [
            "Quarto 1", "Quarto 2"
        ]),
        ("MARÉS", [
            "Quarto 1", "Quarto 2", "Quarto 3"
        ]),
        ("FOZ - Apartamento A", [
            "Quarto 1", "Quarto 2"
        ]),
        ("FOZ - Apartamento B", [
            "Quarto 3", "Quarto 4"
        ]),
        ("DUNAS2", [
            "Quarto 1", "Quarto 2", "Quarto 3"
        ]),
    ]

    # --- Carregar dados de reservas para sugerir saídas ---

    reservas_df = st.session_state.get("reservas_df")
    # Data de referência: checkout mais frequente nos dados inseridos, ou hoje se não houver dados
    def data_referencia_checklist(df=None):
        if df is None or (isinstance(df, pd.DataFrame) and df.empty):
            return date.today()
        try:
            checkouts = pd.to_datetime(df["Check-out"], errors="coerce").dropna()
            if checkouts.empty:
                return date.today()
            return checkouts.dt.date.mode()[0]
        except Exception:
            return date.today()

    hoje = data_referencia_checklist(reservas_df)
    amanha = hoje + timedelta(days=1)
    ontem = hoje - timedelta(days=1)

    # Função para identificar se há saída prevista para o quarto
    import re
    def norm(txt):
        return str(txt).strip().lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")

    def tem_saida_sugerida(alojamento, quarto):
        df = st.session_state.get("reservas_editor_df")
        if df is None or df.empty:
            df = st.session_state.get("reservas_df")
        if df is None or df.empty:
            return False

        quarto_norm = norm(quarto)
        # Alojamento base: "FOZ - Apartamento A" → "foz"
        aloj_checklist = norm(alojamento.split(" - ")[0])

        # Detecta se é uma cama (ex: "Quarto 4 - Cama 1") e extrai o número correcto
        is_cama = "cama" in quarto_norm
        if is_cama:
            m_cama = re.search(r"cama\s*(\d+)", quarto_norm)
            item_num = m_cama.group(1) if m_cama else None
        else:
            m_quarto = re.search(r"(\d+)", quarto_norm)
            item_num = m_quarto.group(1) if m_quarto else None

        def tipo_e_num_unidade(txt):
            """Classifica a unidade da reserva e extrai o número identificador.
            Retorna ('casa', None) | ('cama', str) | ('quarto', str) | ('ambiguo', str) | (None, None)

            Exemplos:
              'quarto duplo nº1 casa de banho privada' → ('quarto', '1')
              'quarto twin nº5 beliche com casa de banho privada' → ('quarto', '5')
              'double room no. 3 private bathroom' → ('quarto', '3')
              'cama 2' / 'bed 2' / 'bunk 2' → ('cama', '2')
              'two bedroom house' / 'casa inteira' → ('casa', None)
              'quarto 1' / 'room 1' → ('quarto', '1')
              '1' / 'q1' → ('ambiguo', '1')
            """
            # Propriedade inteira
            palavras_casa = ["two bedroom", "house", "entire", "completo", "inteiro",
                             "casa inteira", "apartamento inteiro", "whole"]
            if any(p in txt for p in palavras_casa):
                return ('casa', None)

            # Detecta tipo pela presença de palavras-chave
            is_cama_u = bool(re.search(r'\b(?:cama|bed|bunk|beliche)\b', txt, re.IGNORECASE))
            is_quarto_u = bool(re.search(
                r'\b(?:quarto|room|suite|double|twin|single|triple|duplo|individual|triplo|quadruplo)\b',
                txt, re.IGNORECASE
            ))

            # Extrai número: nº/no./nr. têm prioridade, depois palavra+número,
            # depois "- N" no fim (ex: "Quarto Duplo com wc partilhada - 1"), depois standalone
            num = None
            m = re.search(r'\bn[ro]?[º°]?\.?\s*(\d+)\b', txt, re.IGNORECASE)
            if m:
                num = m.group(1)
            else:
                m = re.search(
                    r'\b(?:cama|bed|bunk|room|quarto|suite)\s+(\d+)\b',
                    txt, re.IGNORECASE
                )
                if m:
                    num = m.group(1)
                else:
                    # Número após traço no fim: "Quarto Duplo com wc partilhada - 1"
                    m = re.search(r'-\s*(\d+)\s*$', txt)
                    if m:
                        num = m.group(1)
                    else:
                        # Último número na string: "Quarto Twin 1", "Double Room 2"
                        m = re.search(r'(\d+)\s*$', txt)
                        if m:
                            num = m.group(1)
                        else:
                            m = re.fullmatch(r'([a-z]?)(\d+)', txt.strip())
                            if m:
                                prefix = m.group(1).lower()
                                num = m.group(2)
                                # Prefixo Q = quarto, C = cama (abreviaturas do unidade_curta)
                                if prefix == 'q':
                                    return ('quarto', num)
                                if prefix == 'c':
                                    return ('cama', num)

            # Quarto tem prioridade sobre cama quando ambos os termos estão presentes
            # (ex: "Quarto Twin nº5 beliche" é um quarto com beliches, não uma cama)
            if is_quarto_u:
                return ('quarto', num)
            if is_cama_u:
                return ('cama', num)
            if num:
                return ('ambiguo', num)
            return (None, None)

        for _, row in df.iterrows():
            try:
                checkin = pd.to_datetime(row["Check-in"]).date() if pd.notna(row.get("Check-in")) else None
                checkout = pd.to_datetime(row["Check-out"]).date() if pd.notna(row.get("Check-out")) else None
                if checkout != hoje or checkin is None:
                    continue

                aloj_row = norm(str(row.get("Alojamento", "")))
                # Exige igualdade exata para evitar confusão DUNAS vs DUNAS2
                if aloj_checklist != aloj_row:
                    continue

                unidade_raw = norm(str(row.get("Unidade", "")))

                # Unidade vazia ou igual ao nome do alojamento = propriedade inteira
                if not unidade_raw or unidade_raw == aloj_checklist:
                    return True

                # Suporta unidades múltiplas separadas por vírgula (ex: "Q1, Q4, Q2")
                partes_unidade = [p.strip() for p in unidade_raw.split(",") if p.strip()]
                for unidade in partes_unidade:
                    tipo_u, num_u = tipo_e_num_unidade(unidade)

                    # Propriedade inteira: marca todos os quartos e camas
                    if tipo_u == 'casa':
                        return True

                    if is_cama and item_num:
                        if tipo_u in ('cama', 'ambiguo') and num_u == item_num:
                            return True
                    elif item_num:
                        if tipo_u in ('quarto', 'ambiguo') and (num_u == item_num or num_u is None):
                            return True
            except Exception:
                continue
        return False

    def tem_fica(alojamento, quarto):
        """Retorna True se houver uma reserva com mais de 1 noite que passa esta noite
        (checkin <= hoje < checkout), ou seja, o hóspede fica — não há limpeza."""
        df = st.session_state.get("reservas_editor_df")
        if df is None or df.empty:
            df = st.session_state.get("reservas_df")
        if df is None or df.empty:
            return False

        aloj_checklist = norm(alojamento.split(" - ")[0])
        quarto_norm = norm(quarto)
        is_cama = "cama" in quarto_norm
        if is_cama:
            m_cama = re.search(r"cama\s*(\d+)", quarto_norm)
            item_num = m_cama.group(1) if m_cama else None
        else:
            m_quarto = re.search(r"(\d+)", quarto_norm)
            item_num = m_quarto.group(1) if m_quarto else None

        for _, row in df.iterrows():
            try:
                checkin  = pd.to_datetime(row["Check-in"]).date()  if pd.notna(row.get("Check-in"))  else None
                checkout = pd.to_datetime(row["Check-out"]).date() if pd.notna(row.get("Check-out")) else None
                if checkin is None or checkout is None:
                    continue
                # Hóspede está cá hoje mas não sai hoje (fica = não precisa limpeza)
                if not (checkin <= hoje < checkout):
                    continue

                aloj_row = norm(str(row.get("Alojamento", "")))
                if aloj_checklist not in aloj_row and aloj_row not in aloj_checklist:
                    continue

                unidade_raw = norm(str(row.get("Unidade", "")))

                # Propriedade inteira
                palavras_casa = ["two bedroom", "house", "entire", "completo", "inteiro",
                                 "casa inteira", "apartamento inteiro", "whole"]
                if any(p in unidade_raw for p in palavras_casa):
                    return True
                if not unidade_raw or unidade_raw == aloj_checklist:
                    return True

                def _num(txt):
                    m = re.search(r'\bn[ro]?[º°]?\.?\s*(\d+)\b', txt, re.IGNORECASE)
                    if m: return m.group(1)
                    m = re.search(r'\b(?:cama|bed|bunk|room|quarto|suite)\s+(\d+)\b', txt, re.IGNORECASE)
                    if m: return m.group(1)
                    m = re.search(r'-\s*(\d+)\s*$', txt)
                    if m: return m.group(1)
                    m = re.search(r'(\d+)\s*$', txt)
                    if m: return m.group(1)
                    m = re.fullmatch(r'([a-z]?)(\d+)', txt.strip())
                    if m:
                        return m.group(2)
                    return None

                # Suporta unidades múltiplas separadas por vírgula
                for unidade in [p.strip() for p in unidade_raw.split(",") if p.strip()]:
                    num_u = _num(unidade)
                    # Reconhece prefixos Q/C além de palavras completas
                    _pref = re.fullmatch(r'([a-z]?)(\d+)', unidade.strip())
                    _pref_letter = _pref.group(1).lower() if _pref else ''
                    is_cama_u = bool(re.search(r'\b(?:cama|bed|bunk|beliche)\b', unidade, re.IGNORECASE)) or _pref_letter == 'c'
                    is_quarto_u = bool(re.search(
                        r'\b(?:quarto|room|suite|double|twin|single|triple|duplo|individual|triplo|quadruplo)\b',
                        unidade, re.IGNORECASE)) or _pref_letter == 'q'

                    if is_cama and item_num:
                        if (is_cama_u or not is_quarto_u) and (num_u == item_num or num_u is None):
                            return True
                    elif item_num:
                        if (is_quarto_u or not is_cama_u) and (num_u == item_num or num_u is None):
                            return True
            except Exception:
                continue
        return False

    # Estado das checkboxes — carregado do ficheiro em refresh_data_from_storage; garante existência
    if "saidas_checklist" not in st.session_state:
        _df_for_checklist = st.session_state.get("reservas_editor_df") or st.session_state.get("reservas_df")
        st.session_state["saidas_checklist"] = load_saidas_checklist(_df_for_checklist)

    # Renderização da checklist
    # Cores dos alojamentos (igual ao acesso rápido)
    ALOJAMENTO_EMOJIS = {
        "ABH": "🟥",
        "AFH": "🟦",
        "PIPO": "🟩",
        "DUNAS": "🟨",
        "DUNAS2": "🟪",
        "FOZ": "🟧",
        "ESCAPE": "🟫",
        "MARES": "⬛",
    }
    ALOJAMENTO_BUTTON_COLORS = {
        "ABH": "#ef4444",
        "AFH": "#3b82f6",
        "PIPO": "#22c55e",
        "DUNAS": "#facc15",
        "DUNAS2": "#a855f7",
        "FOZ": "#fb923c",
        "ESCAPE": "#92400e",
        "MARES": "#374151",
    }



    for alojamento, quartos in CHECKLIST_STRUCTURE:
        # Extrai a chave base do alojamento para cor
        if alojamento.startswith("FOZ"):
            aloj_key = "FOZ"
        elif alojamento.startswith("DUNAS2"):
            aloj_key = "DUNAS2"
        elif alojamento.startswith("DUNAS"):
            aloj_key = "DUNAS"
        elif alojamento.startswith("ESCAPE"):
            aloj_key = "ESCAPE"
        elif alojamento.startswith("PIPO"):
            aloj_key = "PIPO"
        elif alojamento.startswith("AFH"):
            aloj_key = "AFH"
        elif alojamento.startswith("ABH"):
            aloj_key = "ABH"
        elif alojamento.startswith("MARÉS") or alojamento.startswith("MARES"):
            aloj_key = "MARES"
        else:
            aloj_key = alojamento.upper()
        cor = ALOJAMENTO_BUTTON_COLORS.get(aloj_key, "#9ca3af")
        st.markdown(f"<div style='background:{cor};color:#fff;padding:0.5rem 1rem;border-radius:10px;margin-top:1.5rem;margin-bottom:0.5rem;display:inline-block;font-weight:700;font-size:1.2rem;'>"
                    f"{alojamento}"
                    f"</div>", unsafe_allow_html=True)
        col1, col2 = st.columns([1, 5])
        # Inicializa o estado ANTES do botão
        # Auto-detecta apenas quando o widget state não existe (após import/reset)
        # Se já existe, respeita o valor manual do utilizador
        for quarto in quartos:
            key = f"saida_{alojamento}_{quarto}"
            if key not in st.session_state:
                sugerida = tem_saida_sugerida(alojamento, quarto)
                st.session_state["saidas_checklist"][key] = bool(sugerida)
                st.session_state[key] = bool(sugerida)
            elif key not in st.session_state["saidas_checklist"]:
                st.session_state["saidas_checklist"][key] = bool(st.session_state[key])

        # Se a flag de marcar todos deste alojamento estiver ativa, marca todos e limpa a flag
        marcar_flag = f"marcar_todos_flag_{alojamento}"
        if st.session_state.get(marcar_flag, False):
            for quarto in quartos:
                key = f"saida_{alojamento}_{quarto}"
                st.session_state["saidas_checklist"][key] = True
                st.session_state[key] = True  # atualiza também o estado do widget diretamente
            st.session_state[marcar_flag] = False

        with col1:
            marcar_todos = st.button("Marcar todos", key=f"marcar_todos_{alojamento}")
            if marcar_todos:
                st.session_state[marcar_flag] = True
                st.rerun()
        with col2:
            changed = False
            for quarto in quartos:
                key = f"saida_{alojamento}_{quarto}"
                fica = tem_fica(alojamento, quarto)
                label = (
                    f'{quarto} \u00a0\u2014\u00a0 <span style="'
                    f'background:#e8f4fd;color:#1a6fa0;font-size:0.8em;font-weight:600;'
                    f'padding:1px 6px;border-radius:4px;border:1px solid #9ecfed;">Fica</span>'
                    if fica else quarto
                )
                new_val = st.checkbox(
                    label if not fica else quarto,
                    value=st.session_state["saidas_checklist"][key],
                    key=key,
                    help="Hóspede fica esta noite — marcar só se necessário limpar." if fica else None,
                )
                if fica:
                    st.caption("Fica esta noite")
                if new_val != st.session_state["saidas_checklist"][key]:
                    st.session_state["saidas_checklist"][key] = new_val
                    changed = True
            if changed:
                save_saidas_checklist(st.session_state["saidas_checklist"], st.session_state.get("reservas_editor_df"))

    st.divider()

uploaded_files = []
all_data = []
import_submit = False

with tab_importar:
    uploaded_files = st.file_uploader(
        "Carregar ficheiros do Booking",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files:
        for file in uploaded_files:
            df = pd.read_excel(file)

            st.write(f"Ficheiro: {file.name}")

            suggested_aloj = suggest_alojamento_from_filename(file.name)
            suggested_idx = ALOJAMENTOS.index(suggested_aloj) if suggested_aloj in ALOJAMENTOS else 0

            alojamento = st.selectbox(
                f"Alojamento para {file.name}",
                ALOJAMENTOS,
                index=suggested_idx,
                key=file.name
            )

            try:
                def limpar_unidade(texto):
                    import re
                    if pd.isna(texto):
                        return texto
                    # Mantém o conteúdo entre parênteses (pode conter info útil como "twin").
                    texto = re.sub(r"[()]", "", str(texto))
                    # Normaliza espaços múltiplos e vírgulas duplicadas
                    texto = re.sub(r",\s*,", ",", texto)
                    texto = re.sub(r"\s*;\s*", ", ", texto)
                    texto = re.sub(r",\s*$", "", texto.strip())
                    texto = re.sub(r"\s+", " ", texto)
                    return texto.strip()

                def _find_col_idx(df_in, patterns, default_idx=None):
                    import re

                    for i, col_name in enumerate(df_in.columns):
                        col_txt = str(col_name).strip().lower()
                        if any(re.search(p, col_txt, flags=re.IGNORECASE) for p in patterns):
                            return i

                    if default_idx is not None and default_idx < len(df_in.columns):
                        return default_idx
                    return None

                def _find_people_col_idx(df_in, default_idx=None):
                    import re

                    cols = list(df_in.columns)

                    # 1) Prioriza cabeçalhos exatos mais comuns para contagem de pessoas.
                    exact_patterns = [
                        r"^people$",
                        r"^pessoas?$",
                        r"^guests?$",
                        r"^occupancy$",
                        r"^pax$",
                        r"^n[úu]mero\s+de\s+pessoas$",
                    ]
                    for i, col_name in enumerate(cols):
                        col_txt = str(col_name).strip().lower()
                        if any(re.search(p, col_txt, flags=re.IGNORECASE) for p in exact_patterns):
                            return i

                    # 2) Caso não seja exato, escolhe o candidato mais "numérico" e evita colunas de nome/titular.
                    include_patterns = [r"\bpeople\b", r"\bpessoas?\b", r"\bguests?\b", r"\boccup", r"\bpax\b", r"\badults?\b"]
                    exclude_patterns = [r"name", r"nome", r"titular", r"holder", r"cliente", r"guest\s*name"]

                    best_idx = None
                    best_score = -1.0

                    for i, col_name in enumerate(cols):
                        col_txt = str(col_name).strip().lower()
                        if any(re.search(p, col_txt, flags=re.IGNORECASE) for p in exclude_patterns):
                            continue
                        if not any(re.search(p, col_txt, flags=re.IGNORECASE) for p in include_patterns):
                            continue

                        series = pd.to_numeric(df_in.iloc[:, i], errors="coerce")
                        score = float(series.notna().mean()) if len(series) > 0 else 0.0
                        if score > best_score:
                            best_score = score
                            best_idx = i

                    if best_idx is not None:
                        return best_idx

                    return None

                nome_idx = _find_col_idx(df, [r"\bnome\b", r"guest", r"h[oó]spede", r"cliente"], default_idx=2)
                checkin_idx = _find_col_idx(df, [r"check\s*[-_ ]?in", r"entrada"], default_idx=3)
                checkout_idx = _find_col_idx(df, [r"check\s*[-_ ]?out", r"sa[ií]da"], default_idx=4)
                pessoas_idx = _find_people_col_idx(df)
                unidade_idx = _find_col_idx(
                    df,
                    [r"\bunit\s*type\b", r"\bunit\b", r"\bunidade\b", r"\bquarto\b", r"\broom\b", r"apart", r"accomm", r"\bcama\b", r"\bbed\b"],
                    default_idx=22,
                )

                unidade_series = (
                    df.iloc[:, unidade_idx].apply(limpar_unidade)
                    if unidade_idx is not None
                    else pd.Series([None] * len(df), index=df.index)
                )

                pessoas_series = (
                    df.iloc[:, pessoas_idx].apply(normalize_pessoas_value)
                    if pessoas_idx is not None
                    else pd.Series([1] * len(df), index=df.index)
                )

                df_clean = pd.DataFrame({
                    "Nome": df.iloc[:, nome_idx] if nome_idx is not None else "",
                    "Check-in": df.iloc[:, checkin_idx] if checkin_idx is not None else pd.NaT,
                    "Check-out": df.iloc[:, checkout_idx] if checkout_idx is not None else pd.NaT,
                    "Pessoas": pessoas_series,
                    "Unidade": unidade_series,
                    "Alojamento": normalize_alojamento(alojamento),
                    "Origem": "Importada",
                })
                all_data.append(df_clean)
            except Exception as e:
                show_pink_alert(f"Erro no ficheiro {file.name}: {e}")

        import_submit = st.button("Importar reservas", type="primary")


with tab_inserir:
    st.subheader("Adicionar reserva direta")

    _df_sug = st.session_state.get("reservas_df", pd.DataFrame()).copy()
    checkin_options = _existing_date_options(_df_sug, "Check-in")
    checkout_options = _existing_date_options(_df_sug, "Check-out")

    with st.form("manual_reserva_form", clear_on_submit=True):
        mf1, mf2, mf3 = st.columns(3)
        with mf1:
            manual_nome = st.text_input("Nome do hóspede")
        with mf2:
            manual_alojamento = st.selectbox("Alojamento", ALOJAMENTOS, key="manual_alojamento")
        with mf3:
            manual_unidade = st.text_input("Unidade (ex: Quarto 1)")

        mf4, mf5, mf6 = st.columns(3)
        with mf4:
            manual_checkin = _choose_date_with_suggestions(
                "Check-in",
                checkin_options,
                datetime.now().date(),
                key_prefix="manual_checkin",
            )
        with mf5:
            manual_checkout = _choose_date_with_suggestions(
                "Check-out",
                checkout_options,
                manual_checkin,
                key_prefix="manual_checkout",
            )
        with mf6:
            st.markdown('<div style="height:52px"></div>', unsafe_allow_html=True)
            manual_pessoas = st.number_input("Pessoas", min_value=1, step=1, value=1)

        mf7, mf8, mf9 = st.columns(3)
        with mf7:
            manual_hora_pa = st.selectbox(
                "Hora PA",
                options=["nenhuma"] + build_suggested_times(),
                index=0,
                help="Hora do pequeno-almoço.",
            )
        with mf8:
            manual_pa_pago = st.selectbox(
                "PA pago",
                options=["Sim", "Não"],
                index=1,
                help="Por defeito fica 'Não'.",
            )
        with mf9:
            manual_notas = st.text_input("Notas", placeholder="Escreve uma nota...")

        manual_submit = st.form_submit_button("Adicionar hóspede")

    if manual_submit:
        _nome_txt = str(manual_nome).strip()
        _unidade_txt = str(manual_unidade).strip()
        if not _nome_txt:
            show_pink_alert("Indica o nome do hóspede para adicionar a reserva direta.")
        elif manual_checkout < manual_checkin:
            show_pink_alert("A data de check-out não pode ser anterior ao check-in.")
        else:
            _manual_reserva = {
                "Nome": _nome_txt,
                "Check-in": manual_checkin,
                "Check-out": manual_checkout,
                "Pessoas": int(manual_pessoas),
                "Unidade": _unidade_txt,
                "Alojamento": normalize_alojamento(manual_alojamento),
                "Origem": "Direta",
                "Hora PA": None if not manual_hora_pa or manual_hora_pa == "nenhuma" else manual_hora_pa,
                "PA pago": "Sim" if manual_pa_pago == "Sim" else None,
                "Notas": str(manual_notas).strip() if str(manual_notas).strip() else None,
            }
            _base_df = st.session_state.get("reservas_df", pd.DataFrame()).copy()
            if _is_duplicate_reserva(_base_df, _manual_reserva):
                show_pink_alert(
                    "Já existe um hóspede com os mesmos dados (nome, datas, pessoas, unidade e alojamento)."
                )
            else:
                _manual_df = pd.DataFrame([_manual_reserva])
                _merged_df, _ = merge_new_reservas(_base_df, _manual_df)
                _merged_df = sanitize_optional_columns(_merged_df)
                st.session_state["reservas_df"] = _merged_df
                st.session_state["reservas_editor_df"] = _merged_df
                save_reservas(_merged_df)
                st.session_state["saidas_checklist"] = load_saidas_checklist(_merged_df)
                st.success("Reserva direta adicionada com sucesso.")
                st.rerun()

    st.divider()
    st.subheader("Quartos disponíveis hoje")

    if "qf_quantidade" not in st.session_state:
        st.session_state["qf_quantidade"] = 1
    if st.session_state.pop("qf_quantidade_reset", False):
        st.session_state["qf_quantidade"] = 1

    with st.form("quartos_form", clear_on_submit=True):
        qf1, qf2, qf3 = st.columns(3)
        with qf1:
            qf_alojamento = st.selectbox("Alojamento", ALOJAMENTOS)
        with qf2:
            _unidade_opcoes = [f"Quarto {i}" for i in range(1, 7)] + [f"Cama {i}" for i in range(1, 7)]
            qf_unidade = st.selectbox("Unidade", _unidade_opcoes, help="Camas disponíveis apenas no ABH")
        with qf3:
            qf_preco = st.number_input("Preço (€)", min_value=0.0, step=5.0, format="%.0f")
        qf4, qf5, qf6 = st.columns(3)
        with qf4:
            qf_pessoas = st.number_input("Pessoas (opcional)", min_value=0, step=1, value=0)
        with qf5:
            qf_notas = st.text_input("Notas (opcional)", placeholder="ex: casa de banho partilhada")
        with qf6:
            qf_quantidade = st.number_input(
                "Quantidade",
                min_value=1,
                step=1,
                key="qf_quantidade",
                help="Se Unidade for Quarto/Cama com número, adiciona em sequência.",
            )
        qf_submit = st.form_submit_button("Adicionar quarto(s) disponível(eis)")

    if qf_submit:
        unidades_para_adicionar = expand_unidade_sequence(qf_unidade, qf_quantidade)
        existentes = {
            (str(item.get("alojamento", "")).strip().upper(), str(item.get("unidade", "")).strip().lower())
            for item in st.session_state["quartos_disponiveis"]
        }

        added_count = 0
        skipped_count = 0
        for unidade_item in unidades_para_adicionar:
            key_unidade = (str(qf_alojamento).strip().upper(), str(unidade_item).strip().lower())
            if key_unidade in existentes:
                skipped_count += 1
                continue

            st.session_state["quartos_disponiveis"].append({
                "alojamento": qf_alojamento,
                "unidade": unidade_item,
                "preco": float(qf_preco),
                "pessoas": int(qf_pessoas) if qf_pessoas > 0 else None,
                "notas": str(qf_notas).strip() or None,
            })
            existentes.add(key_unidade)
            added_count += 1

        save_quartos(st.session_state["quartos_disponiveis"])
        if added_count > 0 and skipped_count == 0:
            st.success(f"{added_count} quarto(s) adicionado(s) em {qf_alojamento}.")
        elif added_count > 0 and skipped_count > 0:
            st.warning(f"{added_count} adicionado(s), {skipped_count} já existiam.")
        else:
            st.info("Nenhum quarto adicionado (já existiam todos).")
        st.session_state["qf_quantidade_reset"] = True
        st.rerun()

    _quartos_lista = st.session_state.get("quartos_disponiveis", [])
    if "inserir_selected_quarto_idx" not in st.session_state:
        st.session_state["inserir_selected_quarto_idx"] = None
    if _quartos_lista:
        for _qi2, _q2 in enumerate(_quartos_lista):
            _b2 = ALOJAMENTO_BADGES.get(_q2["alojamento"].upper(), "⬜")
            _p2 = _q2.get("preco", 0)
            _p2_str = f"{int(_p2)} €" if _p2 == int(_p2) else f"{_p2:.2f} €"
            _pax2 = f" · {_q2['pessoas']} pax" if _q2.get("pessoas") else ""
            _nota2 = f" · {_q2['notas']}" if _q2.get("notas") else ""
            _is_selected = st.session_state["inserir_selected_quarto_idx"] == _qi2
            _row_txt = f"{_b2} {_q2['alojamento']} — {_q2['unidade']} — {_p2_str}{_pax2}{_nota2}"
            _qcol1, _qcol2 = st.columns([4.5, 1])
            with _qcol1:
                if st.button(
                    f"▾ {_row_txt}" if _is_selected else _row_txt,
                    key=f"sel_quarto_{_qi2}",
                    use_container_width=True,
                ):
                    st.session_state["inserir_selected_quarto_idx"] = None if _is_selected else _qi2
                    st.rerun()
            with _qcol2:
                if st.button("", key=f"quick_rm_quarto_{_qi2}", use_container_width=True, help="Remover quarto"):
                    st.session_state["quartos_disponiveis"].pop(_qi2)
                    save_quartos(st.session_state["quartos_disponiveis"])
                    st.session_state["inserir_selected_quarto_idx"] = None
                    st.rerun()

            if st.session_state["inserir_selected_quarto_idx"] == _qi2:
                with st.form(key=f"edit_quarto_form_{_qi2}"):
                    eq1, eq2, eq3 = st.columns(3)
                    with eq1:
                        edit_unidade_q = st.text_input("Unidade", value=str(_q2.get("unidade", "") or ""))
                    with eq2:
                        edit_preco_q = st.number_input(
                            "Preço (€)",
                            min_value=0.0,
                            step=5.0,
                            value=float(_q2.get("preco", 0.0) or 0.0),
                            format="%.0f",
                        )
                    with eq3:
                        edit_pessoas_q = st.number_input(
                            "Pessoas (opcional)",
                            min_value=0,
                            step=1,
                            value=int(_q2.get("pessoas") or 0),
                        )

                    edit_notas_q = st.text_input("Notas (opcional)", value=str(_q2.get("notas", "") or ""))

                    eb1, eb2 = st.columns(2)
                    with eb1:
                        save_edit_q = st.form_submit_button("Guardar alterações")
                    with eb2:
                        cancel_edit_q = st.form_submit_button("Cancelar")

                if save_edit_q:
                    nova_unidade_q = str(edit_unidade_q).strip()
                    if not nova_unidade_q:
                        show_pink_alert("A unidade não pode ficar vazia.")
                    else:
                        st.session_state["quartos_disponiveis"][_qi2] = {
                            "alojamento": _q2.get("alojamento"),
                            "unidade": nova_unidade_q,
                            "preco": float(edit_preco_q),
                            "pessoas": int(edit_pessoas_q) if edit_pessoas_q > 0 else None,
                            "notas": str(edit_notas_q).strip() or None,
                        }
                        save_quartos(st.session_state["quartos_disponiveis"])
                        st.session_state["inserir_selected_quarto_idx"] = None
                        st.success("Quarto atualizado com sucesso.")
                        st.rerun()

                if cancel_edit_q:
                    st.session_state["inserir_selected_quarto_idx"] = None
                    st.rerun()
    else:
        st.caption("Sem quartos disponíveis registados.")

df_guardado = st.session_state["reservas_df"].copy()
df_final = pd.DataFrame()
novas_reservas_count = 0
reservas_atualizadas_count = 0

if import_submit and all_data:
    df_importado = pd.concat(all_data, ignore_index=True)
    df_importado = normalize_pessoas_column(df_importado)
    df_final, novas_reservas_count, reservas_atualizadas_count = merge_imported_reservas(df_guardado, df_importado)
    df_final = sanitize_optional_columns(df_final)
    df_final = normalize_pessoas_column(df_final)
    st.session_state["reservas_df"] = df_final
    save_reservas(df_final)
    # Reset completo após importação — limpa widget states para forçar re-detecção
    for _k in [k for k in st.session_state if isinstance(k, str) and k.startswith("saida_")]:
        del st.session_state[_k]
    st.session_state["saidas_checklist"] = {}
    _import_conflicts = detect_conflicts(df_final)
    if _import_conflicts:
        st.warning(f"**{len(_import_conflicts)} conflito(s) detectado(s) após importação:**")
        for _c in _import_conflicts:
            st.markdown(_c)
elif not df_guardado.empty:
    df_final = df_guardado.copy()

if not df_final.empty:
    df_final = sanitize_optional_columns(df_final)
    df_final = normalize_pessoas_column(df_final)
    if "Origem" not in df_final.columns:
        # Reservas antigas (sem origem) passam a importadas para permitir limpeza seletiva.
        df_final["Origem"] = "Importada"
    else:
        df_final["Origem"] = (
            df_final["Origem"]
            .astype(str)
            .str.strip()
            .replace({"": "Importada", "none": "Importada", "nan": "Importada", "None": "Importada"})
        )
    suggested_times = build_suggested_times()

    if "Hora PA" not in df_final.columns:
        df_final["Hora PA"] = None
    if "PA pago" not in df_final.columns:
        df_final["PA pago"] = None
    if "Notas" not in df_final.columns:
        df_final["Notas"] = None

    ordered_cols = [c for c in DISPLAY_COL_ORDER if c in df_final.columns]
    extra_cols = [c for c in df_final.columns if c not in ordered_cols]
    df_final = df_final[ordered_cols + extra_cols]

    st.session_state["reservas_editor_df"] = sanitize_optional_columns(df_final.copy())


    with tab_acesso_rapido:
        render_quick_access_tab(st.session_state["reservas_editor_df"], suggested_times)

    with tab_reservas:
        render_reservas_editor(suggested_times)


    with tab_notas:
        # Garante que df_pa está sempre definido
        if "reservas_editor_df" in st.session_state:
            edited_df = sanitize_optional_columns(st.session_state["reservas_editor_df"].copy())
            df_pa = edited_df.copy()
            df_pa = df_pa[df_pa["Hora PA"].notna() & (df_pa["Hora PA"] != "")]
        else:
            df_pa = pd.DataFrame()

        def unidade_curta(valor_unidade):
            import re
            if pd.isna(valor_unidade) or not str(valor_unidade).strip():
                return ""
            partes = [p.strip() for p in str(valor_unidade).split(",") if p.strip()]
            resultado = []
            for parte in partes:
                m_quarto = re.search(r"quarto[^\d]*(\d+)", parte, flags=re.IGNORECASE)
                if m_quarto:
                    resultado.append(f"Q{int(m_quarto.group(1))}")
                    continue
                m_cama = re.search(r"cama[^\d]*(\d+)", parte, flags=re.IGNORECASE)
                if m_cama:
                    resultado.append(f"C{int(m_cama.group(1))}")
                    continue
                resultado.append(parte)
            return ", ".join(resultado)

        st.subheader(" Notas gerais")
        st.caption("Estas notas entram no fim do texto exportado de pequenos-almoços e não substituem as notas de cada reserva.")
        st.text_area(
            "Notas para exportação",
            key="notas_gerais_pa_editor",
            placeholder="Ex.: Avisar cozinha sobre alergias gerais do dia...",
            height=180,
        )
        if st.button("Guardar notas", key="save_notas_gerais_btn", type="primary"):
            notas_limpas = str(st.session_state.get("notas_gerais_pa_editor", "")).strip()
            st.session_state["notas_gerais_pa"] = notas_limpas
            save_notas_gerais(notas_limpas)
            st.success("Notas gerais guardadas.")

        # --- Transfers ---
        st.divider()
        st.subheader("Transfers")
        transfers = st.session_state.get("transfers", [])
        editing_idx = st.session_state.get("editing_transfer_idx", None)

        if transfers:
            for i, t in enumerate(transfers):
                label = f"**{t['nome']}** — {t['alojamento']}"
                if t.get("unidade"):
                    label += f" ({unidade_curta(t['unidade'])})"
                if editing_idx == i:
                    st.markdown(label)
                    novo_texto = st.text_area("Editar transfer", value=t["texto"], key=f"edit_transfer_texto_{i}")
                    c1, c2 = st.columns([1, 1])
                    with c1:
                        if st.button("Guardar", key=f"save_transfer_{i}", type="primary"):
                            transfers[i]["texto"] = novo_texto.strip()
                            st.session_state["transfers"] = transfers
                            st.session_state["editing_transfer_idx"] = None
                            save_transfers(transfers)
                            st.rerun()
                    with c2:
                        if st.button("Cancelar", key=f"cancel_transfer_{i}"):
                            st.session_state["editing_transfer_idx"] = None
                            st.rerun()
                else:
                    col_t, col_edit, col_del = st.columns([5, 1, 1])
                    with col_t:
                        st.markdown(f"{label}  \n{t['texto']}")
                    with col_edit:
                        if st.button("Editar", key=f"edit_transfer_{i}"):
                            st.session_state["editing_transfer_idx"] = i
                            st.rerun()
                    with col_del:
                        if st.button("Remover", key=f"remove_transfer_{i}"):
                            transfers.pop(i)
                            st.session_state["transfers"] = transfers
                            st.session_state["editing_transfer_idx"] = None
                            save_transfers(transfers)
                            st.rerun()
            st.divider()

        reservas_para_transfer = st.session_state.get("reservas_editor_df", pd.DataFrame())
        if not reservas_para_transfer.empty:
            opcoes_transfer = [""] + [
                f"{r['Nome']} — {r['Alojamento']}" + (f" ({unidade_curta(r.get('Unidade',''))})" if r.get('Unidade') else "") + f"  [{pd.to_datetime(r.get('Check-in'), errors='coerce').strftime('%d/%m') if pd.notna(pd.to_datetime(r.get('Check-in'), errors='coerce')) else ''}]"
                for _, r in reservas_para_transfer.iterrows()
            ]
            with st.form("form_add_transfer", clear_on_submit=True):
                sel = st.selectbox("Reserva", opcoes_transfer, key="transfer_sel_reserva")
                texto_transfer = st.text_area("Informação do transfer", key="transfer_texto", placeholder="Ex.: Transfer para aeroporto às 10h00")
                if st.form_submit_button("Adicionar transfer", type="primary"):
                    if sel and texto_transfer.strip():
                        idx = opcoes_transfer.index(sel) - 1
                        row = reservas_para_transfer.iloc[idx]
                        novo = {
                            "nome": row["Nome"],
                            "alojamento": row["Alojamento"],
                            "unidade": str(row.get("Unidade", "") or ""),
                            "checkin": str(row.get("Check-in", "") or ""),
                            "texto": texto_transfer.strip(),
                        }
                        transfers.append(novo)
                        st.session_state["transfers"] = transfers
                        save_transfers(transfers)
                        st.rerun()
        else:
            st.caption("Importa reservas para poder associar transfers.")

        # --- Lista Pequenos-Almoços (movida do tab_pa) ---
        st.divider()
        st.subheader(" Exportar lista")

        def gerar_lista(df_pa, saidas_checklist=None, checklist_structure=None, bold_asterisk=False, transfers=None, df_reservas=None):
            def bold(x):
                return f"*{x}*" if bold_asterisk else f"**{x}**"
            def nota_valida(v):
                return pd.notna(v) and str(v).strip() and str(v).strip().lower() not in {"none", "nan", "nat"}

            # --- Pequenos-Almoços ---
            linhas = [bold('PEQUENOS-ALMOÇOS:')]
            for hora in sorted(df_pa["Hora PA"].unique()):
                grupo = df_pa[df_pa["Hora PA"] == hora]
                linhas.append(f"{bold(f'{hora}h')}")
                for _, r in grupo.iterrows():
                    nome = r["Nome"]
                    aloj = r["Alojamento"]
                    unidade = unidade_curta(r["Unidade"])
                    pax_num = pd.to_numeric(r.get("Pessoas"), errors="coerce")
                    pax = int(pax_num) if pd.notna(pax_num) else 0
                    is_pago = str(r.get("PA pago", "")).strip().lower() in ["sim", "yes", "s"]
                    nota = r.get("Notas", None)
                    nota_texto = str(nota).strip() if nota_valida(nota) else ""
                    tags = []
                    if is_pago:
                        tags.append("pago")
                    if nota_texto:
                        tags.append(nota_texto)
                    sufixo = f" ({'; '.join(tags)})" if tags else ""
                    linhas.append(f"{nome}  {aloj} {unidade} - {pax} pax{sufixo}")
            total_pax = total_pessoas_col(df_pa)
            linhas.append(bold(f"Total de pessoas: {total_pax}"))

            # --- Saídas ---
            # saidas_checklist reflecte o estado actual das checkboxes (auto + manual)
            # fica é calculado directamente dos dados para não depender de estado externo
            if checklist_structure:
                df_res = df_reservas if df_reservas is not None else pd.DataFrame()
                ref_hoje = None
                if not df_res.empty:
                    try:
                        cos = pd.to_datetime(df_res["Check-out"], errors="coerce").dropna()
                        if not cos.empty:
                            ref_hoje = cos.dt.date.mode()[0]
                    except Exception:
                        pass

                def _fica(alojamento, quarto):
                    """True se há hóspede com estada que continua hoje."""
                    if df_res.empty or ref_hoje is None:
                        return False
                    aloj_norm = str(alojamento).split(" - ")[0].strip().lower()
                    q_norm = str(quarto).strip().lower()
                    is_cama_q = "cama" in q_norm
                    m_num = re.search(r'(\d+)', q_norm)
                    item_num = m_num.group(1) if m_num else None
                    for _, row in df_res.iterrows():
                        try:
                            checkin = pd.to_datetime(row["Check-in"]).date() if pd.notna(row.get("Check-in")) else None
                            checkout = pd.to_datetime(row["Check-out"]).date() if pd.notna(row.get("Check-out")) else None
                            if not checkin or not checkout or not (checkin <= ref_hoje < checkout):
                                continue
                            if str(row.get("Alojamento", "")).strip().lower() != aloj_norm:
                                continue
                            unidade_raw = str(row.get("Unidade", "") or "").strip().lower()
                            if not unidade_raw or unidade_raw == aloj_norm:
                                return True
                            for parte in unidade_raw.split(","):
                                parte = parte.strip()
                                is_cama_u = bool(re.search(r'\b(?:cama|bed|bunk|beliche)\b', parte, re.IGNORECASE))
                                is_quarto_u = bool(re.search(r'\b(?:quarto|room|suite|double|twin|single|duplo)\b', parte, re.IGNORECASE))
                                m_u = (re.search(r'\bnr?[º°]?\.?\s*(\d+)\b', parte, re.IGNORECASE) or
                                       re.search(r'\b(?:cama|quarto|room)\s+(\d+)\b', parte, re.IGNORECASE) or
                                       re.search(r'(\d+)\s*$', parte))
                                num_u = m_u.group(1) if m_u else None
                                if is_cama_q:
                                    if (is_cama_u or not is_quarto_u) and (num_u == item_num or num_u is None):
                                        return True
                                elif item_num:
                                    if (is_quarto_u or not is_cama_u) and (num_u == item_num or num_u is None):
                                        return True
                        except Exception:
                            continue
                    return False

                linhas_saidas = []
                for alojamento, quartos in checklist_structure:
                    quartos_checked = [q for q in quartos if saidas_checklist and saidas_checklist.get(f"saida_{alojamento}_{q}")]
                    quartos_fica = [q for q in quartos if _fica(alojamento, q)]
                    if not quartos_checked and not quartos_fica:
                        continue
                    saidas_sem_fica = [q for q in quartos_checked if q not in quartos_fica]
                    partes = []
                    if saidas_sem_fica:
                        if len(saidas_sem_fica) == len(quartos) - len(quartos_fica):
                            partes.append(bold("COMPLETO"))
                        else:
                            partes.append(", ".join(saidas_sem_fica))
                    if quartos_fica:
                        partes.append(f"Fica: {', '.join(quartos_fica)}")
                    if partes:
                        linhas_saidas.append(f"{bold(alojamento)} {' | '.join(partes)}")
                if linhas_saidas:
                    linhas.append("")
                    linhas.append(bold("SAÍDAS:"))
                    linhas.extend(linhas_saidas)

            # --- Transfers ---
            transfers_lista = transfers or []
            if transfers_lista:
                linhas.append("")
                linhas.append(bold("TRANSFERS:"))
                for t in transfers_lista:
                    unidade_t = unidade_curta(t.get("unidade", "")) if t.get("unidade") else ""
                    ref = f"{t['nome']} — {t['alojamento']}" + (f" ({unidade_t})" if unidade_t else "")
                    linhas.append(f"{ref}: {t['texto']}")

            # --- Notas ---
            notas_gerais = str(st.session_state.get("notas_gerais_pa", "")).strip()
            if notas_gerais:
                linhas.append("")
                linhas.append(bold("NOTAS:"))
                linhas.append(notas_gerais)
            return "\n".join(linhas)


        def gerar_lista_md(df_pa, saidas_checklist=None, checklist_structure=None, transfers=None):
            return gerar_lista(df_pa, saidas_checklist, checklist_structure, bold_asterisk=False, transfers=transfers).replace("*", "**")

        if st.button("Gerar lista em texto", key="btn_gerar_lista_md_notas"):
            checklist_structure = CHECKLIST_STRUCTURE if 'CHECKLIST_STRUCTURE' in locals() or 'CHECKLIST_STRUCTURE' in globals() else None
            saidas_checklist = {}
            if checklist_structure:
                for _aloj, _qs in checklist_structure:
                    for _q in _qs:
                        _k = f"saida_{_aloj}_{_q}"
                        saidas_checklist[_k] = bool(st.session_state.get(_k, False))
            transfers_export = st.session_state.get("transfers", [])
            _df1 = st.session_state.get("reservas_editor_df")
            df_res_export = _df1 if (_df1 is not None and not _df1.empty) else st.session_state.get("reservas_df")
            lista_texto = gerar_lista(df_pa, saidas_checklist, checklist_structure, bold_asterisk=True, transfers=transfers_export, df_reservas=df_res_export)
            st.code(lista_texto, language=None)


    edited_df = sanitize_optional_columns(st.session_state["reservas_editor_df"].copy())

    df_pa = edited_df.copy()
    df_pa = df_pa[df_pa["Hora PA"].notna() & (df_pa["Hora PA"] != "")]
    with tab_pa:
        reservas_df = st.session_state.get("reservas_df")
        if reservas_df is not None and not reservas_df.empty and len(df_pa) > 0:
            df_pa = df_pa.sort_values("Hora PA")
            df_pa_resumo = df_pa.copy()
            df_pa_resumo["Pessoas"] = pd.to_numeric(df_pa_resumo["Pessoas"], errors="coerce").fillna(0)
            resumo = df_pa_resumo.groupby("Hora PA")["Pessoas"].sum().reset_index()
            total_pa_dia = total_pessoas_col(df_pa)

            st.subheader(" Resumo por Hora")
            for _, row in resumo.iterrows():
                hora = row["Hora PA"]
                total_num = pd.to_numeric(row.get("Pessoas"), errors="coerce")
                total = int(total_num) if pd.notna(total_num) else 0
                if total >= 16:
                    show_pink_alert(f"{hora} → {total} pessoas")
                else:
                    st.success(f"{hora} → {total} pessoas")
                grupo = df_pa[df_pa["Hora PA"] == hora]
                st.dataframe(grupo[["Nome", "Alojamento", "Unidade", "Pessoas", "PA pago"]].fillna(""), width='stretch')

            st.divider()

            st.subheader(" Ocupação do Espaço (Pequeno-Almoço)")

            occupation_data = build_occupation_data(df_pa, suggested_times)

            df_occupation = pd.DataFrame(occupation_data)
            df_occupation["Verde (<16)"] = df_occupation["Pessoas"].where(df_occupation["Pessoas"] < 16, 0)
            df_occupation["Vermelho (>=16)"] = df_occupation["Pessoas"].where(df_occupation["Pessoas"] >= 16, 0)

            chart_long = df_occupation.melt(
                id_vars=["Hora"],
                value_vars=["Verde (<16)", "Vermelho (>=16)"],
                var_name="Faixa",
                value_name="Total",
            )
            occupation_chart = (
                alt.Chart(chart_long)
                .mark_bar()
                .encode(
                    x=alt.X("Hora:N", sort=suggested_times, title="Hora"),
                    y=alt.Y("Total:Q", stack="zero", scale=alt.Scale(domain=[0, 20]), title="Pessoas"),
                    color=alt.Color(
                        "Faixa:N",
                        scale=alt.Scale(
                            domain=["Verde (<16)", "Vermelho (>=16)"],
                            range=[THEME["chart_ok"], THEME["chart_over"]],
                        ),
                        legend=alt.Legend(title=None),
                    ),
                    tooltip=["Hora:N", "Faixa:N", "Total:Q"],
                )
                .properties(height=280)
            )
            st.altair_chart(occupation_chart, use_container_width=True)

            for row in occupation_data:
                if row["Pessoas"] >= 16:
                    pessoas_num = pd.to_numeric(row.get("Pessoas"), errors="coerce")
                    pessoas_total = int(pessoas_num) if pd.notna(pessoas_num) else 0
                    show_pink_alert(f"{row['Hora']} → {pessoas_total} pessoas")

            st.subheader(" Total de Pequenos-Almoços (dia)")
            st.metric("Total de pessoas", total_pa_dia)

            st.divider()
        else:
            reservas_df = st.session_state.get("reservas_df")
            if reservas_df is None or reservas_df.empty:
                st.info("Sem dados ainda. Importa ficheiros no separador 'Importar' para começar.")
            else:
                st.info("Nenhuma reserva com hora de pequeno-almoço definida ainda.")

with tab_importar:
    st.divider()
    st.subheader("Limpar dados da aplicação")
    st.warning("⚠ Esta ação apaga todos os dados: reservas, quartos disponíveis, notas gerais, checklist de saídas e estado de sessão. Não pode ser desfeita.")

    if not st.session_state.get("limpar_confirmar_pendente", False):
        if st.button("Limpar TUDO", key="limpar_tudo_btn"):
            st.session_state["limpar_confirmar_pendente"] = True
            st.rerun()
    else:
        st.markdown(
            f'<style>'
            f'div[data-testid="stButton"] button[kind="primary"]#limpar_confirmar_btn,'
            f'div[data-testid="stButton"]:has(button[key="limpar_confirmar_btn"]) button {{'
            f'  background-color: {THEME["tab_active"]} !important;'
            f'  border-color: {THEME["tab_active"]} !important;'
            f'  color: #fff !important;'
            f'}}'
            f'</style>',
            unsafe_allow_html=True,
        )
        col_sim, col_nao, _ = st.columns([1, 1, 4])
        with col_sim:
            if st.button("Confirmar limpeza", key="limpar_confirmar_btn", type="primary"):
                st.session_state["limpar_confirmar_pendente"] = False
                st.session_state["reservas_df"] = pd.DataFrame()
                st.session_state["reservas_editor_df"] = pd.DataFrame()
                st.session_state["quartos_disponiveis"] = []
                st.session_state["notas_gerais_pa"] = ""
                st.session_state.pop("notas_gerais_pa_editor", None)
                st.session_state["saidas_checklist"] = {}
                st.session_state["transfers"] = []
                for k in list(st.session_state.keys()):
                    if k.startswith("saida_") or k.startswith("marcar_todos_flag_"):
                        del st.session_state[k]
                for f in [RESERVAS_FILE, QUARTOS_FILE, NOTAS_GERAIS_FILE, SAIDAS_FILE, TRANSFERS_FILE]:
                    if f.exists():
                        f.unlink()
                if USE_SUPABASE and _supabase_client:
                    try:
                        _supabase_client.table("reservas").delete().eq("id", 1).execute()
                        _supabase_client.table("reservas").delete().eq("id", 2).execute()
                    except Exception:
                        pass
                st.rerun()
        with col_nao:
            if st.button("Cancelar", key="limpar_cancelar_btn"):
                st.session_state["limpar_confirmar_pendente"] = False
                st.rerun()