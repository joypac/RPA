import streamlit as st
import streamlit as st

st.set_page_config(
    page_title="Gestão de Reservas",
    page_icon="🍳"
)


import pandas as pd
import altair as alt
from datetime import datetime, timedelta
import json
import html
from pathlib import Path

st.set_page_config(layout="wide")

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


ALOJAMENTO_BADGES = {
    "ABH": "🟥",
    "AFH": "🟦",
    "PIPO": "🟩",
    "DUNAS": "🟨",
    "DUNAS2": "🟪",
    "FOZ": "🟧",
    "ESCAPE": "🟫",
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

tab_reservas, tab_pa, tab_importar, tab_guardar = st.tabs(
    ["Reservas", "Pequenos almoços", "Importar", "Limpar"]
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
                ["ABH", "AFH", "PIPO", "DUNAS", "DUNAS2", "FOZ", "ESCAPE"],
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
    with tab_reservas:
        st.info("Sem dados ainda. Importa ficheiros no separador 'Importar' para começar.")
    with tab_pa:
        st.info("Sem dados ainda. Importa ficheiros no separador 'Importar' para começar.")
    with tab_guardar:
        st.info("Sem dados para guardar/limpar neste momento.")