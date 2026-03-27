import streamlit as st
import pandas as pd
import altair as alt
from datetime import datetime, timedelta, date
import json
import html
from pathlib import Path

st.set_page_config(
    page_title="Gestão de Reservas",
    page_icon="🌅",
    layout="wide",
)

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

st.markdown(
        f"""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;700&display=swap');

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
        </style>
        """,
        unsafe_allow_html=True,
)

RESERVAS_FILE = Path("reservas.json")
QUARTOS_FILE = Path("quartos_disponiveis.json")
RESERVA_KEY_COLS = ["Nome", "Check-in", "Check-out", "Pessoas", "Unidade", "Alojamento"]
DISPLAY_COL_ORDER = [
    "Alojamento",
    "Nome",
    "Check-in",
    "Check-out",
    "Pessoas",
    "Unidade",
    "Hora PA",
    "PA pago",
    "Notas",
]

ALOJAMENTOS = ["ABH", "AFH", "PIPO", "DUNAS", "DUNAS2", "FOZ", "ESCAPE", "MARES"]

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

    if not RESERVAS_FILE.exists():
        return pd.DataFrame()

    try:
        with RESERVAS_FILE.open("r", encoding="utf-8") as f:
            data = json.load(f)
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
            return pd.DataFrame(data)
    except Exception as e:
        show_pink_alert(f"Erro ao carregar reservas guardadas: {e}")

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

    with RESERVAS_FILE.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    st.session_state["last_saved_at"] = saved_at


def load_quartos():
    if not QUARTOS_FILE.exists():
        return []
    try:
        with QUARTOS_FILE.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def save_quartos(quartos_list):
    with QUARTOS_FILE.open("w", encoding="utf-8") as f:
        json.dump(quartos_list, f, ensure_ascii=False, indent=2)


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
                    occupation += int(person_row["Pessoas"])
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
            messages.append(f"⚠️ {row['Hora']} → {int(row['Pessoas'])} pessoas")
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

    partes = [p.strip() for p in str(unidade_value).split(",") if p.strip()]
    for parte in partes:
        m_quarto = re.search(r"quarto[^\d]*(\d+)", parte, flags=re.IGNORECASE)
        if m_quarto:
            return f"Q{int(m_quarto.group(1))}"

        m_cama = re.search(r"cama[^\d]*(\d+)", parte, flags=re.IGNORECASE)
        if m_cama:
            return f"C{int(m_cama.group(1))}"

    return ""


def format_nome_com_quarto(nome_value, unidade_value):
    nome = "" if pd.isna(nome_value) else str(nome_value).strip()
    room_tag = extract_room_tag(unidade_value)
    if nome and room_tag:
        return f"{nome} ({room_tag})"
    if room_tag:
        return room_tag
    return nome


def format_quartos_text(unidade_value):
    if pd.isna(unidade_value):
        return "Sem quarto definido"

    partes = [p.strip() for p in str(unidade_value).split(",") if p.strip()]
    if not partes:
        return "Sem quarto definido"

    quartos = []
    for parte in partes:
        import re

        m_quarto = re.search(r"quarto[^\d]*(\d+)", parte, flags=re.IGNORECASE)
        if m_quarto:
            quartos.append(f"Quarto {int(m_quarto.group(1))}")
            continue

        m_cama = re.search(r"cama[^\d]*(\d+)", parte, flags=re.IGNORECASE)
        if m_cama:
            quartos.append(f"Cama {int(m_cama.group(1))}")
            continue

        quartos.append(parte)

    return ", ".join(quartos)


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
    alojamentos = (
        quick_df["Alojamento"]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .str.upper()
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
                            _nova = pd.DataFrame([{
                                "Nome": _ins_nome_txt,
                                "Alojamento": _q["alojamento"],
                                "Unidade": _q["unidade"],
                                "Pessoas": int(_ins_pessoas),
                                "Check-in": _today,
                                "Check-out": _tomorrow,
                                "Hora PA": None if _ins_hora == "nenhuma" else _ins_hora,
                                "Notas": str(_ins_notas).strip() or None,
                            }])
                            _merged, _ = merge_new_reservas(_base, _nova)
                            _merged = sanitize_optional_columns(_merged)
                            st.session_state["reservas_df"] = _merged
                            st.session_state["reservas_editor_df"] = _merged
                            st.session_state["current_df"] = _merged
                            save_reservas(_merged)
                            st.session_state["quartos_disponiveis"].pop(_qi)
                            save_quartos(st.session_state["quartos_disponiveis"])
                            st.session_state["quick_inserir_quarto_idx"] = None
                            st.success(f"{_ins_nome_txt} inserido no {_q['unidade']} ({_q['alojamento']}).")
                            st.rerun()
        return

    pessoas_df = quick_df[quick_df["Alojamento"].astype(str).str.upper() == selected_aloj].copy()
    pessoas_df = pessoas_df.reset_index().rename(columns={"index": "_idx"})

    # Passo 2: apenas nomes do alojamento selecionado
    if selected_idx is None:
        ctop1, ctop2 = st.columns([1.5, 4])
        with ctop1:
            if st.button("⬅️ Voltar aos alojamentos", key="quick_back_to_aloj"):
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
        if st.button("⬅️ Voltar aos nomes", key="quick_back_to_names"):
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
    st.success(f"Quarto(s): {quartos_text}")

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
        st.session_state["current_df"] = updated_df.copy()
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

    display_df = editor_df.copy()
    if "Alojamento" in display_df.columns:
        display_df["Alojamento"] = display_df["Alojamento"].apply(format_alojamento_badge)
    if "Nome" in display_df.columns:
        display_df["Nome"] = display_df["Nome"].fillna("")
    display_df["Check-in/Check-out"] = display_df.apply(
        lambda row: format_checkin_checkout(row.get("Check-in"), row.get("Check-out")),
        axis=1,
    )
    display_df = display_df.drop(columns=[c for c in ["Check-in", "Check-out"] if c in display_df.columns])
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
    edited_from_table = st.data_editor(
        display_df,
        key="reservas_editor",
        width='stretch',
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
        st.session_state["current_df"] = updated_df.copy()
        st.session_state["reservas_df"] = updated_df.copy()
        save_reservas(updated_df)
        editor_df = updated_df

    st.divider()

    select_df = editor_df.reset_index().rename(columns={"index": "_idx"})
    reserva_idx = st.selectbox(
        "Selecionar reserva",
        options=select_df["_idx"].tolist(),
        format_func=lambda i: (
            f"{'' if pd.isna(select_df.loc[select_df['_idx'] == i, 'Nome'].iloc[0]) else select_df.loc[select_df['_idx'] == i, 'Nome'].iloc[0]} | "
            f"{'' if pd.isna(select_df.loc[select_df['_idx'] == i, 'Alojamento'].iloc[0]) else select_df.loc[select_df['_idx'] == i, 'Alojamento'].iloc[0]} "
            f"{'' if pd.isna(select_df.loc[select_df['_idx'] == i, 'Unidade'].iloc[0]) else select_df.loc[select_df['_idx'] == i, 'Unidade'].iloc[0]}"
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

    c1, c2, c3 = st.columns([1, 1, 1.5])
    with c1:
        nova_hora = st.selectbox(
            "Hora PA",
            options=hora_options,
            index=hora_options.index(hora_default) if hora_default is not None else None,
            placeholder="nenhuma",
            key=f"hora_pa_update_{reserva_idx}",
        )
    with c2:
        novo_pago = st.selectbox(
            "PA pago",
            options=pago_options,
            index=pago_options.index(pago_default) if pago_default is not None else None,
            placeholder="Não",
            key=f"pa_pago_update_{reserva_idx}",
        )
    with c3:
        novas_notas = st.text_input(
            "Notas",
            value=notas_default,
            key=f"notas_update_{reserva_idx}",
            placeholder="Escreve uma nota...",
        )

    if st.button("Aplicar alterações à reserva", key="apply_reserva_update"):
        df_before_update = editor_df.copy()
        editor_df.loc[reserva_idx, "Hora PA"] = None if not nova_hora or nova_hora == "nenhuma" else nova_hora
        editor_df.loc[reserva_idx, "PA pago"] = None if not novo_pago or novo_pago == "Não" else novo_pago
        editor_df.loc[reserva_idx, "Notas"] = None if not str(novas_notas).strip() or str(novas_notas).strip().lower() == "nenhuma" else novas_notas
        editor_df = sanitize_optional_columns(editor_df)

        old_overcrowding = set(build_overcrowding_messages(df_before_update, suggested_times))
        new_overcrowding = set(build_overcrowding_messages(editor_df, suggested_times))
        created_overcrowding = sorted(new_overcrowding - old_overcrowding)
        if created_overcrowding:
            st.session_state["pending_overcrowding_messages"] = created_overcrowding
            st.session_state["show_overcrowding_ack"] = True

        st.session_state["reservas_editor_df"] = editor_df.copy()
        st.session_state["current_df"] = editor_df.copy()
        st.session_state["reservas_df"] = editor_df.copy()
        save_reservas(editor_df)
        st.success("Reserva atualizada.")
        st.rerun()

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
                checkin_atual_date = checkin_atual.date() if pd.notna(checkin_atual) else None
                if checkin_edit_options:
                    ci_idx = checkin_edit_options.index(checkin_atual_date) if checkin_atual_date in checkin_edit_options else 0
                    edit_checkin = st.selectbox(
                        "Check-in",
                        options=checkin_edit_options,
                        index=ci_idx,
                        format_func=lambda d: d.strftime("%d/%m/%Y"),
                    )
                else:
                    edit_checkin = st.date_input("Check-in", value=checkin_atual_date)
            with ef5:
                checkout_atual = pd.to_datetime(row_edit.get("Check-out"), errors="coerce")
                checkout_atual_date = checkout_atual.date() if pd.notna(checkout_atual) else None
                if checkout_edit_options:
                    co_idx = checkout_edit_options.index(checkout_atual_date) if checkout_atual_date in checkout_edit_options else 0
                    edit_checkout = st.selectbox(
                        "Check-out",
                        options=checkout_edit_options,
                        index=co_idx,
                        format_func=lambda d: d.strftime("%d/%m/%Y"),
                    )
                else:
                    edit_checkout = st.date_input("Check-out", value=checkout_atual_date)
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
            else:
                editor_df.loc[reserva_idx, "Nome"] = nome_edit_txt
                editor_df.loc[reserva_idx, "Alojamento"] = edit_alojamento
                editor_df.loc[reserva_idx, "Pessoas"] = int(edit_pessoas)
                editor_df.loc[reserva_idx, "Check-in"] = edit_checkin
                editor_df.loc[reserva_idx, "Check-out"] = edit_checkout
                editor_df.loc[reserva_idx, "Unidade"] = str(edit_unidade).strip() or None
                editor_df.loc[reserva_idx, "Hora PA"] = None if edit_hora_pa == "nenhuma" else edit_hora_pa
                editor_df.loc[reserva_idx, "PA pago"] = "Sim" if edit_pa_pago == "Sim" else None
                editor_df.loc[reserva_idx, "Notas"] = str(edit_notas).strip() or None
                editor_df = sanitize_optional_columns(editor_df)
                st.session_state["reservas_editor_df"] = editor_df.copy()
                st.session_state["current_df"] = editor_df.copy()
                st.session_state["reservas_df"] = editor_df.copy()
                save_reservas(editor_df)
                st.success("Reserva editada com sucesso.")
                st.rerun()

    st.divider()
    st.subheader("🛏️ Ocupação desta noite")
    st.metric("Total de hóspedes", total_hospedes_esta_noite(editor_df))

    st.divider()
    st.caption("Exportação")

    from io import BytesIO

    excel_buffer = BytesIO()
    display_df_export = sanitize_optional_columns(editor_df).fillna("")
    display_df_export.to_excel(excel_buffer, index=False, engine="openpyxl", sheet_name="Reservas")
    excel_buffer.seek(0)
    st.download_button(
        label="📥 Exportar para Excel",
        data=excel_buffer.getvalue(),
        file_name="reservas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


render_reservas_editor = (
    st.fragment(_render_reservas_editor_impl)
    if hasattr(st, "fragment")
    else _render_reservas_editor_impl
)


if "current_df" not in st.session_state:
    st.session_state["current_df"] = pd.DataFrame()
if "reservas_df" not in st.session_state:
    st.session_state["reservas_df"] = load_reservas()
if "pending_overcrowding_messages" not in st.session_state:
    st.session_state["pending_overcrowding_messages"] = []
if "show_overcrowding_ack" not in st.session_state:
    st.session_state["show_overcrowding_ack"] = False
if "quartos_disponiveis" not in st.session_state:
    st.session_state["quartos_disponiveis"] = load_quartos()

st.markdown("### Gestão de Reservas + Pequenos-Almoços")

st.caption(f"Guardado pela última vez: {get_last_saved_text()}")

if st.session_state["show_overcrowding_ack"] and st.session_state["pending_overcrowding_messages"]:
    show_pink_alert("Foram criados horários com sobrelotação. Confirma para continuar.")
    for msg in st.session_state["pending_overcrowding_messages"]:
        show_pink_alert(msg)
    if st.button("OK", key="ack_overcrowding"):
        st.session_state["show_overcrowding_ack"] = False
        st.session_state["pending_overcrowding_messages"] = []
        st.rerun()

tab_acesso_rapido, tab_reservas, tab_pa, tab_importar, tab_inserir, tab_guardar = st.tabs(
    ["Acesso rápido", "Reservas", "Pequenos almoços", "Importar", "Inserir", "Limpar"]
)

uploaded_files = []
all_data = []

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

            alojamento = st.selectbox(
                f"Alojamento para {file.name}",
                ALOJAMENTOS,
                key=file.name
            )

            try:
                def limpar_unidade(texto):
                    import re
                    if pd.isna(texto):
                        return texto
                    # Remove tudo o que estiver entre parênteses (inclusive os parênteses)
                    texto = re.sub(r"\s*\([^)]*\)", "", str(texto))
                    # Normaliza espaços múltiplos e vírgulas duplicadas
                    texto = re.sub(r",\s*,", ",", texto)
                    texto = re.sub(r",\s*$", "", texto.strip())
                    return texto.strip()

                df_clean = pd.DataFrame({
                    "Nome": df.iloc[:, 2],
                    "Check-in": df.iloc[:, 3],
                    "Check-out": df.iloc[:, 4],
                    "Pessoas": df.iloc[:, 8],
                    "Unidade": df.iloc[:, 22].apply(limpar_unidade),
                    "Alojamento": alojamento
                })
                all_data.append(df_clean)
            except Exception as e:
                show_pink_alert(f"Erro no ficheiro {file.name}: {e}")


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
            manual_pessoas = st.number_input("Pessoas", min_value=1, step=1, value=1)

        mf4, mf5, mf6 = st.columns(3)
        with mf4:
            if checkin_options:
                manual_checkin = st.selectbox(
                    "Check-in",
                    options=checkin_options,
                    format_func=lambda d: d.strftime("%d/%m/%Y"),
                    help="Datas sugeridas com base nos outros hóspedes.",
                )
            else:
                manual_checkin = st.date_input("Check-in")
        with mf5:
            if checkout_options:
                manual_checkout = st.selectbox(
                    "Check-out",
                    options=checkout_options,
                    format_func=lambda d: d.strftime("%d/%m/%Y"),
                    help="Datas sugeridas com base nos outros hóspedes.",
                )
            else:
                manual_checkout = st.date_input("Check-out")
        with mf6:
            manual_unidade = st.text_input("Unidade (ex: Quarto 1)")

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
                index=0,
                help="Por defeito fica 'Sim' para reservas diretas com pequeno-almoço incluído.",
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
            _manual_df = pd.DataFrame(
                [
                    {
                        "Nome": _nome_txt,
                        "Check-in": manual_checkin,
                        "Check-out": manual_checkout,
                        "Pessoas": int(manual_pessoas),
                        "Unidade": _unidade_txt,
                        "Alojamento": manual_alojamento,
                        "Hora PA": None if not manual_hora_pa or manual_hora_pa == "nenhuma" else manual_hora_pa,
                        "PA pago": "Sim" if manual_pa_pago == "Sim" else None,
                        "Notas": str(manual_notas).strip() if str(manual_notas).strip() else None,
                    }
                ]
            )
            _base_df = st.session_state.get("reservas_df", pd.DataFrame()).copy()
            _merged_df, _ = merge_new_reservas(_base_df, _manual_df)
            _merged_df = sanitize_optional_columns(_merged_df)
            st.session_state["reservas_df"] = _merged_df
            st.session_state["reservas_editor_df"] = _merged_df
            st.session_state["current_df"] = _merged_df
            save_reservas(_merged_df)
            st.success("Reserva direta adicionada com sucesso.")
            st.rerun()

    st.divider()
    st.subheader("Quartos disponíveis hoje")

    with st.form("quartos_form", clear_on_submit=True):
        qf1, qf2, qf3 = st.columns(3)
        with qf1:
            qf_alojamento = st.selectbox("Alojamento", ALOJAMENTOS)
        with qf2:
            _unidade_opcoes = [f"Quarto {i}" for i in range(1, 7)] + [f"Cama {i}" for i in range(1, 7)]
            qf_unidade = st.selectbox("Unidade", _unidade_opcoes, help="Camas disponíveis apenas no ABH")
        with qf3:
            qf_preco = st.number_input("Preço (€)", min_value=0.0, step=5.0, format="%.0f")
        qf4, qf5 = st.columns(2)
        with qf4:
            qf_pessoas = st.number_input("Pessoas (opcional)", min_value=0, step=1, value=0)
        with qf5:
            qf_notas = st.text_input("Notas (opcional)", placeholder="ex: casa de banho partilhada")
        qf_submit = st.form_submit_button("Adicionar quarto disponível")

    if qf_submit:
        st.session_state["quartos_disponiveis"].append({
            "alojamento": qf_alojamento,
            "unidade": qf_unidade,
            "preco": float(qf_preco),
            "pessoas": int(qf_pessoas) if qf_pessoas > 0 else None,
            "notas": str(qf_notas).strip() or None,
        })
        save_quartos(st.session_state["quartos_disponiveis"])
        st.success(f"{qf_alojamento} — {qf_unidade} adicionado.")
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
            if st.button(
                f"▾ {_row_txt}" if _is_selected else _row_txt,
                key=f"sel_quarto_{_qi2}",
                use_container_width=True,
            ):
                st.session_state["inserir_selected_quarto_idx"] = None if _is_selected else _qi2
                st.rerun()

            if st.session_state["inserir_selected_quarto_idx"] == _qi2:
                _r1, _r2 = st.columns([1.2, 1.2])
                with _r1:
                    if st.button("Remover quarto", key=f"rm_quarto_{_qi2}", use_container_width=True):
                        st.session_state["quartos_disponiveis"].pop(_qi2)
                        save_quartos(st.session_state["quartos_disponiveis"])
                        st.session_state["inserir_selected_quarto_idx"] = None
                        st.rerun()
                with _r2:
                    if st.button("Cancelar", key=f"cancel_quarto_{_qi2}", use_container_width=True):
                        st.session_state["inserir_selected_quarto_idx"] = None
                        st.rerun()
    else:
        st.caption("Sem quartos disponíveis registados.")

df_guardado = st.session_state["reservas_df"].copy()
df_final = pd.DataFrame()
novas_reservas_count = 0

if all_data:
    df_importado = pd.concat(all_data, ignore_index=True)
    df_final, novas_reservas_count = merge_new_reservas(df_guardado, df_importado)
elif not df_guardado.empty:
    df_final = df_guardado.copy()

if not df_final.empty:
    df_final = sanitize_optional_columns(df_final)
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

    edited_df = sanitize_optional_columns(st.session_state["reservas_editor_df"].copy())
    st.session_state["current_df"] = edited_df.copy()

    df_pa = edited_df.copy()
    df_pa = df_pa[df_pa["Hora PA"].notna() & (df_pa["Hora PA"] != "")]

    with tab_pa:
        if len(df_pa) > 0:
            df_pa = df_pa.sort_values("Hora PA")
            resumo = df_pa.groupby("Hora PA")["Pessoas"].sum().reset_index()
            total_pa_dia = total_pessoas_col(df_pa)

            st.subheader("📊 Resumo por Hora")
            for _, row in resumo.iterrows():
                hora = row["Hora PA"]
                total = int(row["Pessoas"])
                if total >= 16:
                    show_pink_alert(f"{hora} → {total} pessoas")
                else:
                    st.success(f"{hora} → {total} pessoas")
                grupo = df_pa[df_pa["Hora PA"] == hora]
                st.dataframe(grupo[["Nome", "Alojamento", "Unidade", "Pessoas", "PA pago"]].fillna(""), width='stretch')

            st.divider()

            st.subheader("📈 Ocupação do Espaço (Pequeno-Almoço)")

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
                    show_pink_alert(f"{row['Hora']} → {int(row['Pessoas'])} pessoas")

            st.subheader("👥 Total de Pequenos-Almoços (dia)")
            st.metric("Total de pessoas", total_pa_dia)

            st.divider()
            st.subheader("📋 Lista para Pequenos-Almoços")

            def unidade_curta(valor_unidade):
                import re

                if pd.isna(valor_unidade):
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

            def gerar_lista(df_pa):
                def nota_valida(v):
                    return pd.notna(v) and str(v).strip() and str(v).strip().lower() not in {"none", "nan", "nat"}

                linhas = ["🥐 Pequenos-Almoços\n"]
                for hora in sorted(df_pa["Hora PA"].unique()):
                    grupo = df_pa[df_pa["Hora PA"] == hora]
                    linhas.append(f"{hora}h")
                    for _, r in grupo.iterrows():
                        nome = r["Nome"]
                        aloj = r["Alojamento"]
                        unidade = unidade_curta(r["Unidade"])
                        pax = int(r["Pessoas"])
                        is_pago = str(r.get("PA pago", "")).strip().lower() in ["sim", "yes", "s"]
                        nota = r.get("Notas", None)
                        nota_texto = str(nota).strip() if nota_valida(nota) else ""
                        if is_pago and nota_texto:
                            sufixo = f" (pago; {nota_texto})"
                        elif is_pago:
                            sufixo = " (pago)"
                        elif nota_texto:
                            sufixo = f" ({nota_texto})"
                        else:
                            sufixo = ""
                        linhas.append(f"{nome}  {aloj} {unidade} - {pax} pax{sufixo}")
                    linhas.append("")

                linhas.append(f"Total de pessoas: {total_pessoas_col(df_pa)}")

                return "\n".join(linhas)

            def gerar_lista_md(df_pa):
                def nota_valida(v):
                    return pd.notna(v) and str(v).strip() and str(v).strip().lower() not in {"none", "nan", "nat"}

                linhas = ["**🥐 Pequenos-Almoços**\n"]
                for hora in sorted(df_pa["Hora PA"].unique()):
                    grupo = df_pa[df_pa["Hora PA"] == hora]
                    linhas.append(f"**{hora}h**")
                    for _, r in grupo.iterrows():
                        nome = r["Nome"]
                        aloj = r["Alojamento"]
                        unidade = unidade_curta(r["Unidade"])
                        pax = int(r["Pessoas"])
                        is_pago = str(r.get("PA pago", "")).strip().lower() in ["sim", "yes", "s"]
                        nota = r.get("Notas", None)
                        nota_texto = str(nota).strip() if nota_valida(nota) else ""
                        if is_pago and nota_texto:
                            sufixo = f" (pago; {nota_texto})"
                        elif is_pago:
                            sufixo = " (pago)"
                        elif nota_texto:
                            sufixo = f" ({nota_texto})"
                        else:
                            sufixo = ""
                        linhas.append(f"{nome}  {aloj} {unidade} - {pax} pax{sufixo}")
                    linhas.append("")

                linhas.append(f"**Total de pessoas:** {total_pessoas_col(df_pa)}")

                return "\n\n".join(linhas)

            lista_texto = gerar_lista(df_pa)

            col1, col2 = st.columns(2)
            with col1:
                if st.button("📄 Gerar lista em texto"):
                    st.markdown(gerar_lista_md(df_pa))
                    st.code(lista_texto, language=None)
            with col2:
                st.download_button(
                    label="⬇️ Descarregar lista (.txt)",
                    data=lista_texto,
                    file_name="pequenos_almocoslist.txt",
                    mime="text/plain"
                )
        else:
            st.info("Nenhuma reserva com hora de pequeno-almoço definida ainda.")

    with tab_importar:
        if all_data:
            if novas_reservas_count > 0:
                st.success(f"Foram adicionadas {novas_reservas_count} reserva(s) nova(s).")
            else:
                st.info("Não foram encontradas reservas novas no ficheiro importado.")

    with tab_guardar:
        if "confirm_clear" not in st.session_state:
            st.session_state["confirm_clear"] = False

        if st.button("Limpar tudo"):
            st.session_state["confirm_clear"] = True

        if st.session_state["confirm_clear"]:
            st.warning("Confirma que queres apagar todas as reservas guardadas?")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("Confirmar apagar tudo", type="primary"):
                    if USE_SUPABASE:
                        save_reservas(pd.DataFrame())
                    else:
                        st.info("Supabase não ativo. O botão limpar apaga apenas dados da nuvem.")
                    st.session_state["current_df"] = pd.DataFrame()
                    st.session_state["reservas_df"] = pd.DataFrame()
                    st.session_state["pending_overcrowding_messages"] = []
                    st.session_state["show_overcrowding_ack"] = False
                    st.session_state["confirm_clear"] = False
                    st.success("Reservas apagadas com sucesso.")
                    st.rerun()
            with c2:
                if st.button("Cancelar"):
                    st.session_state["confirm_clear"] = False
else:
    with tab_acesso_rapido:
        st.info("Sem dados ainda. Importa ficheiros no separador 'Importar' para começar.")
    with tab_reservas:
        st.info("Sem dados ainda. Importa ficheiros no separador 'Importar' para começar.")
    with tab_pa:
        st.info("Sem dados ainda. Importa ficheiros no separador 'Importar' para começar.")
    with tab_guardar:
        st.info("Sem dados para guardar/limpar neste momento.")