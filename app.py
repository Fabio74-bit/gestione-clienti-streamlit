import streamlit as st
import pandas as pd
from datetime import date, timedelta, datetime
import io
import requests

# --------------------------------------------------------
# CONFIGURAZIONE BASE DELL'APP
# --------------------------------------------------------
st.set_page_config(page_title="Gestione Clienti - Dashboard", layout="wide")

# 🔗 Collegamento PWA (per icona su iPad)
st.markdown(
    '<link rel="manifest" href="manifest.json">',
    unsafe_allow_html=True
)

st.title("📒 Gestione Clienti — Dashboard completa")

st.markdown("""
Questa app mostra i dati dei clienti aggiornati **automaticamente da OneDrive** 📂  
Ogni giorno alle **12:00** viene scaricata una nuova versione del file Excel `GESTIONE_CLIENTI.xlsm`.
""")

# --------------------------------------------------------
# FUNZIONE PER SCARICARE IL FILE DA ONEDRIVE
# --------------------------------------------------------
@st.cache_data(ttl=3600)
def load_excel_from_onedrive():
    """Scarica e carica il file Excel dal link OneDrive (in formato download diretto)."""
    try:
        url = st.secrets["general"]["ONEDRIVE_URL"]
        response = requests.get(url)
        if response.status_code != 200:
            raise Exception(f"Errore nel download del file (HTTP {response.status_code})")

        file_bytes = io.BytesIO(response.content)
        raw = pd.read_excel(file_bytes, sheet_name="Indice", header=None, engine="openpyxl")

        header = raw.iloc[1].tolist()
        data = raw.iloc[2:].copy()
        data.columns = header

        wanted_cols = [
            "Cliente",
            "Ultimo Recall",
            "Ultima Visita",
            "Prossima Scadenza Noleggio",
            "Tot. Contratti (aperti)",
            "TMK",
        ]
        present = [c for c in wanted_cols if c in data.columns]
        df = data[present].copy()
        df = df.dropna(how="all")

        for c in ["Ultimo Recall", "Ultima Visita", "Prossima Scadenza Noleggio"]:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], errors="coerce").dt.date

        if "Tot. Contratti (aperti)" in df.columns:
            df["Tot. Contratti (aperti)"] = pd.to_numeric(df["Tot. Contratti (aperti)"], errors="coerce")

        if "Cliente" in df.columns:
            df = df[df["Cliente"].notna()]

        return df

    except Exception as e:
        st.error(f"❌ Errore nel caricamento automatico da OneDrive: {e}")
        return None


# --------------------------------------------------------
# CONTROLLO ORARIO AGGIORNAMENTO
# --------------------------------------------------------
def should_refresh_data():
    """Controlla se sono passate le 12:00 di oggi: se sì, forza l’aggiornamento."""
    now = datetime.now()
    refresh_time = datetime.combine(date.today(), datetime.strptime("12:00", "%H:%M").time())
    return now >= refresh_time


# --------------------------------------------------------
# CARICAMENTO AUTOMATICO O MANUALE
# --------------------------------------------------------
st.sidebar.header("⚙️ Aggiornamento Dati")
if st.sidebar.button("🔄 Aggiorna ora"):
    st.cache_data.clear()
    st.session_state["last_refresh"] = datetime.now()
    st.toast("✅ Dati aggiornati manualmente!", icon="🔁")

# Carica dati
if "last_refresh" not in st.session_state or should_refresh_data():
    df = load_excel_from_onedrive()
    st.session_state["last_refresh"] = datetime.now()
else:
    df = load_excel_from_onedrive()

if df is None or df.empty:
    st.warning("⚠️ Non sono riuscito a caricare il file da OneDrive. Verifica il link nei secrets.")
    st.stop()


# --------------------------------------------------------
# FILTRI
# --------------------------------------------------------
st.sidebar.header("🔎 Filtri")

query = st.sidebar.text_input("Cerca cliente (nome parziale)", "")

tmk_values = sorted([x for x in df["TMK"].dropna().unique()]) if "TMK" in df.columns else []
tmk_sel = st.sidebar.multiselect("TMK", tmk_values, default=tmk_values if tmk_values else None)

days = st.sidebar.slider("Mostra scadenze entro (giorni)", min_value=0, max_value=365, value=60, step=15)

today = date.today()
deadline_limit = today + timedelta(days=days)

filtered = df.copy()

if query:
    filtered = filtered[filtered["Cliente"].astype(str).str.contains(query, case=False, na=False)]

if "TMK" in filtered.columns and tmk_sel:
    filtered = filtered[filtered["TMK"].isin(tmk_sel)]

if "Prossima Scadenza Noleggio" in filtered.columns:
    mask = (filtered["Prossima Scadenza Noleggio"].isna()) | (filtered["Prossima Scadenza Noleggio"] <= deadline_limit)
    filtered = filtered[mask]


# --------------------------------------------------------
# KPI
# --------------------------------------------------------
k1, k2, k3 = st.columns(3)

with k1:
    st.metric("Clienti mostrati", len(filtered))

with k2:
    tot_aperti = filtered["Tot. Contratti (aperti)"].sum() if "Tot. Contratti (aperti)" in filtered.columns else 0
    st.metric("Totale contratti aperti (mostrati)", int(tot_aperti) if pd.notna(tot_aperti) else 0)

with k3:
    if "Prossima Scadenza Noleggio" in filtered.columns:
        count_scadenze = filtered["Prossima Scadenza Noleggio"].apply(lambda d: pd.notna(d) and d <= deadline_limit).sum()
        st.metric(f"Scadenze entro {days} giorni", int(count_scadenze))
    else:
        st.metric("Scadenze entro X giorni", 0)


# --------------------------------------------------------
# TABELLA
# --------------------------------------------------------
st.subheader("📋 Elenco clienti (dopo i filtri)")
st.dataframe(filtered, use_container_width=True)

csv = filtered.to_csv(index=False).encode("utf-8")
st.download_button("📥 Scarica come CSV", csv, file_name="clienti_filtrati.csv", mime="text/csv")


# --------------------------------------------------------
# GRAFICI
# --------------------------------------------------------
st.subheader("📊 Grafici")

chart_type = st.selectbox(
    "Scegli un grafico",
    ["Clienti per TMK", "Scadenze per mese", "Distribuzione contratti aperti"]
)

if chart_type == "Clienti per TMK" and "TMK" in filtered.columns:
    st.bar_chart(filtered["TMK"].value_counts())

elif chart_type == "Scadenze per mese" and "Prossima Scadenza Noleggio" in filtered.columns:
    tmp = filtered.dropna(subset=["Prossima Scadenza Noleggio"]).copy()
    if not tmp.empty:
        tmp["Mese"] = pd.to_datetime(tmp["Prossima Scadenza Noleggio"]).dt.to_period("M").astype(str)
        st.bar_chart(tmp["Mese"].value_counts().sort_index())
    else:
        st.info("Nessuna scadenza disponibile nei dati filtrati.")

elif chart_type == "Distribuzione contratti aperti" and "Tot. Contratti (aperti)" in filtered.columns:
    st.bar_chart(filtered["Tot. Contratti (aperti)"].fillna(0).astype(int).value_counts().sort_index())


# --------------------------------------------------------
# INFO AGGIUNTIVE
# --------------------------------------------------------
st.caption(f"🕒 Ultimo aggiornamento: {st.session_state['last_refresh'].strftime('%d/%m/%Y %H:%M:%S')}")
st.caption("💡 Il file Excel viene letto automaticamente ogni giorno alle 12:00 da OneDrive.")
