import streamlit as st
import pandas as pd
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
        </style>
        """,
        unsafe_allow_html=True,
)

RESERVAS_FILE = Path("reservas.json")
RESERVA_KEY_COLS = ["Nome", "Check-in", "Check-out", "Pessoas", "Unidade", "Alojamento"]
DISPLAY_COL_ORDER = [
    "Nome",
    "Check-in",
    "Check-out",
    "Pessoas",
    "Unidade",
    "Alojamento",
    "Hora PA",
    "PA pago",
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
                if isinstance(data, list):
                    return pd.DataFrame(data)
        except Exception as e:
            show_pink_alert(f"Erro ao carregar do Supabase: {e}")
        return pd.DataFrame()

    if not RESERVAS_FILE.exists():
        return pd.DataFrame()

    try:
        with RESERVAS_FILE.open("r", encoding="utf-8") as f:
            data = json.load(f)
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

    if USE_SUPABASE:
        try:
            _supabase_client.table("reservas").upsert({"id": 1, "data": records}).execute()
            st.session_state["last_saved_at"] = datetime.now()
            return
        except Exception as e:
            show_pink_alert(f"Erro ao guardar no Supabase: {e}")

    with RESERVAS_FILE.open("w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)


def get_last_saved_text():
    if USE_SUPABASE:
        last = st.session_state.get("last_saved_at")
        if last:
            return last.strftime("%d/%m/%Y %H:%M:%S")
        return "Ainda não guardado"
    if not RESERVAS_FILE.exists():
        return "Ainda não guardado"
    ts = datetime.fromtimestamp(RESERVAS_FILE.stat().st_mtime)
    return ts.strftime("%d/%m/%Y %H:%M:%S")


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
    for col in ("Hora PA", "PA pago"):
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


if "current_df" not in st.session_state:
    st.session_state["current_df"] = pd.DataFrame()
if "reservas_df" not in st.session_state:
    st.session_state["reservas_df"] = load_reservas()
if "pending_overcrowding_messages" not in st.session_state:
    st.session_state["pending_overcrowding_messages"] = []
if "show_overcrowding_ack" not in st.session_state:
    st.session_state["show_overcrowding_ack"] = False

top_title_col, top_save_col = st.columns([4, 1])
with top_title_col:
    st.markdown("### Gestao de Reservas + Pequenos-Almocos")
with top_save_col:
    st.markdown("<div style='height: 2.2rem;'></div>", unsafe_allow_html=True)
    if st.button("💾 Guardar Agora", use_container_width=True):
        current_df = st.session_state.get("current_df", pd.DataFrame())
        if isinstance(current_df, pd.DataFrame) and not current_df.empty:
            save_reservas(current_df)
            st.success("Guardado.")
        else:
            st.warning("Sem dados para guardar.")

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
    ["Reservas", "Pequenos almoços", "Importar", "Guardar e Limpar"]
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
                df_clean = pd.DataFrame({
                    "Nome": df.iloc[:, 2],
                    "Check-in": df.iloc[:, 3],
                    "Check-out": df.iloc[:, 4],
                    "Pessoas": df.iloc[:, 8],
                    "Unidade": df.iloc[:, 22],
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

    ordered_cols = [c for c in DISPLAY_COL_ORDER if c in df_final.columns]
    extra_cols = [c for c in df_final.columns if c not in ordered_cols]
    df_final = df_final[ordered_cols + extra_cols]

    with tab_reservas:
        display_df = sanitize_optional_columns(df_final).fillna("")
        edited_from_table = st.data_editor(
            display_df,
            width='stretch',
            disabled=["Nome", "Check-in", "Check-out", "Pessoas", "Unidade", "Alojamento"],
            column_config={
                "Hora PA": st.column_config.SelectboxColumn(
                    "Hora PA",
                    options=suggested_times,
                    required=False,
                ),
                "PA pago": st.column_config.SelectboxColumn(
                    "PA pago",
                    options=["Sim"],
                    required=False,
                ),
            }
        )

        # Mantém edição direta apenas nas colunas permitidas.
        df_final["Hora PA"] = edited_from_table["Hora PA"]
        df_final["PA pago"] = edited_from_table["PA pago"]
        df_final = sanitize_optional_columns(df_final)

        old_overcrowding = set(build_overcrowding_messages(df_guardado, suggested_times))
        new_overcrowding = set(build_overcrowding_messages(df_final, suggested_times))
        created_overcrowding = sorted(new_overcrowding - old_overcrowding)
        if created_overcrowding:
            st.session_state["pending_overcrowding_messages"] = created_overcrowding
            st.session_state["show_overcrowding_ack"] = True

        st.divider()

        select_df = df_final.reset_index().rename(columns={"index": "_idx"})
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

        current_hora = df_final.loc[reserva_idx, "Hora PA"]
        current_pago = df_final.loc[reserva_idx, "PA pago"]

        hora_options = [""] + suggested_times
        pago_options = ["", "Sim"]

        hora_default = str(current_hora) if pd.notna(current_hora) else ""
        if hora_default not in hora_options:
            hora_default = ""

        pago_default = str(current_pago) if pd.notna(current_pago) else ""
        if pago_default not in pago_options:
            pago_default = ""

        c1, c2 = st.columns(2)
        with c1:
            nova_hora = st.selectbox(
                "Hora PA",
                options=hora_options,
                index=hora_options.index(hora_default),
                key=f"hora_pa_update_{reserva_idx}",
            )
        with c2:
            novo_pago = st.selectbox(
                "PA pago",
                options=pago_options,
                index=pago_options.index(pago_default),
                key=f"pa_pago_update_{reserva_idx}",
            )

        if st.button("Aplicar alterações à reserva", key="apply_reserva_update"):
            df_before_update = df_final.copy()
            df_final.loc[reserva_idx, "Hora PA"] = nova_hora if nova_hora else None
            df_final.loc[reserva_idx, "PA pago"] = novo_pago if novo_pago else None
            df_final = sanitize_optional_columns(df_final)

            old_overcrowding = set(build_overcrowding_messages(df_before_update, suggested_times))
            new_overcrowding = set(build_overcrowding_messages(df_final, suggested_times))
            created_overcrowding = sorted(new_overcrowding - old_overcrowding)
            if created_overcrowding:
                st.session_state["pending_overcrowding_messages"] = created_overcrowding
                st.session_state["show_overcrowding_ack"] = True

            st.session_state["current_df"] = df_final.copy()
            st.session_state["reservas_df"] = df_final.copy()
            save_reservas(df_final)
            st.success("Reserva atualizada.")
            st.rerun()

    edited_df = sanitize_optional_columns(df_final)

    st.session_state["current_df"] = edited_df.copy()
    st.session_state["reservas_df"] = edited_df.copy()
    save_reservas(edited_df)

    df_pa = edited_df.copy()
    df_pa = df_pa[df_pa["Hora PA"].notna() & (df_pa["Hora PA"] != "")]

    with tab_pa:
        if len(df_pa) > 0:
            df_pa = df_pa.sort_values("Hora PA")
            resumo = df_pa.groupby("Hora PA")["Pessoas"].sum().reset_index()

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
            chart_df = df_occupation.set_index("Hora")[["Verde (<16)", "Vermelho (>=16)"]]
            st.bar_chart(chart_df, color=[THEME["chart_ok"], THEME["chart_over"]])

            for row in occupation_data:
                if row["Pessoas"] >= 16:
                    show_pink_alert(f"{row['Hora']} → {int(row['Pessoas'])} pessoas")

            st.divider()
            st.subheader("📋 Lista para Pequenos-Almoços")

            def gerar_lista(df_pa):
                linhas = ["🥐 Pequenos-Almoços\n"]
                for hora in sorted(df_pa["Hora PA"].unique()):
                    grupo = df_pa[df_pa["Hora PA"] == hora]
                    linhas.append(f"{hora}h")
                    for _, r in grupo.iterrows():
                        nome = r["Nome"]
                        aloj = r["Alojamento"]
                        unidade = r["Unidade"]
                        pax = int(r["Pessoas"])
                        pago = " (pago)" if str(r.get("PA pago", "")).strip().lower() in ["sim", "yes", "s"] else ""
                        linhas.append(f"{nome}  {aloj} {unidade} - {pax} pax{pago}")
                    linhas.append("")
                return "\n".join(linhas)

            lista_texto = gerar_lista(df_pa)

            col1, col2 = st.columns(2)
            with col1:
                if st.button("📄 Gerar lista em texto"):
                    st.text_area("Lista para imprimir / copiar", lista_texto, height=400)
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
        st.caption("O botão de guardar principal está no topo e fica sempre visível.")
        if st.button("Guardar agora (neste separador)"):
            save_reservas(edited_df)
            st.success("Reservas guardadas com sucesso.")

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