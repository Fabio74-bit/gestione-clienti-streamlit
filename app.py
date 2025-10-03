
import streamlit as st
import pandas as pd
from datetime import date, timedelta

st.set_page_config(page_title="Gestione Clienti - Dashboard", layout="wide")

st.title("📒 Gestione Clienti — Dashboard semplice")

st.markdown("""
Questa piccola app ti permette di **vedere, filtrare e cercare** i clienti
partendo dal file Excel `GESTIONE_CLIENTI.xlsm` (foglio **Indice**).
Se non hai esperienza: tranquillo, basta seguire i passi qui sotto. 😊
""")

# ---------------------------
# 1) Caricamento dei dati
# ---------------------------
def load_from_excel(file):
    # leggiamo solo il foglio "Indice"
    raw = pd.read_excel(file, sheet_name="Indice", header=None, engine="openpyxl")
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
    # tieni solo righe non vuote
    df = df.dropna(how="all")
    # parse date
    for c in ["Ultimo Recall", "Ultima Visita", "Prossima Scadenza Noleggio"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
    # numeri
    if "Tot. Contratti (aperti)" in df.columns:
        df["Tot. Contratti (aperti)"] = pd.to_numeric(df["Tot. Contratti (aperti)"], errors="coerce")
    # togli righe senza nome cliente
    if "Cliente" in df.columns:
        df = df[df["Cliente"].notna()]
    return df

default_path = "GESTIONE_CLIENTI.xlsm"
uploaded = None

col1, col2 = st.columns([1,2])
with col1:
    use_default = st.checkbox("Usa il file presente nella stessa cartella dell'app (GESTIONE_CLIENTI.xlsm)", value=True)
with col2:
    uploaded = st.file_uploader("Oppure carica il file Excel (.xlsm / .xlsx)", type=["xlsm","xlsx"], accept_multiple_files=False)

df = None
load_error = None

try:
    if use_default:
        df = load_from_excel(default_path)
    elif uploaded is not None:
        df = load_from_excel(uploaded)
except Exception as e:
    load_error = str(e)

if df is None:
    if load_error:
        st.error("Non sono riuscito a leggere il file. Errore: " + load_error)
    st.info("Carica un file .xlsm/.xlsx o spunta 'Usa il file...' sopra.")
    st.stop()

# ---------------------------
# 2) Filtri semplici
# ---------------------------
st.sidebar.header("🔎 Filtri")

# Cerca per nome cliente
query = st.sidebar.text_input("Cerca cliente (nome parziale)", "")

# Filtro TMK (assegnatario) se presente
tmk_values = sorted([x for x in df["TMK"].dropna().unique()]) if "TMK" in df.columns else []
tmk_sel = st.sidebar.multiselect("TMK", tmk_values, default=tmk_values if tmk_values else None)

# Filtro "scadenze entro X giorni"
days = st.sidebar.slider("Mostra scadenze entro (giorni)", min_value=0, max_value=365, value=60, step=15)

today = date.today()
deadline_limit = today + timedelta(days=days)

# Applichiamo i filtri
filtered = df.copy()

if query:
    filtered = filtered[filtered["Cliente"].astype(str).str.contains(query, case=False, na=False)]

if "TMK" in filtered.columns and tmk_sel:
    filtered = filtered[filtered["TMK"].isin(tmk_sel)]

if "Prossima Scadenza Noleggio" in filtered.columns:
    # Tieni le righe con scadenza nulla oppure entro la data limite
    mask = (filtered["Prossima Scadenza Noleggio"].isna()) | (filtered["Prossima Scadenza Noleggio"] <= deadline_limit)
    filtered = filtered[mask]

# ---------------------------
# 3) KPI e tabella
# ---------------------------
k1, k2, k3 = st.columns(3)

with k1:
    st.metric("Clienti mostrati", len(filtered))

with k2:
    tot_aperti = filtered["Tot. Contratti (aperti)"].sum() if "Tot. Contratti (aperti)" in filtered.columns else 0
    st.metric("Totale contratti aperti (mostrati)", int(tot_aperti) if pd.notna(tot_aperti) else 0)

with k3:
    # quante scadenze entro X giorni
    if "Prossima Scadenza Noleggio" in filtered.columns:
        count_scadenze = filtered["Prossima Scadenza Noleggio"].apply(lambda d: pd.notna(d) and d <= deadline_limit).sum()
        st.metric(f"Scadenze entro {days} giorni", int(count_scadenze))
    else:
        st.metric("Scadenze entro X giorni", 0)

st.subheader("📋 Elenco clienti (dopo i filtri)")
st.dataframe(filtered, use_container_width=True)

# Download CSV
csv = filtered.to_csv(index=False).encode("utf-8")
st.download_button("Scarica come CSV", csv, file_name="clienti_filtrati.csv", mime="text/csv")

# ---------------------------
# 4) Grafici veloci
# ---------------------------
st.subheader("📊 Grafici")

chart_type = st.selectbox("Scegli un grafico", ["Clienti per TMK", "Scadenze per mese", "Distribuzione contratti aperti"])

if chart_type == "Clienti per TMK" and "TMK" in filtered.columns:
    st.bar_chart(filtered["TMK"].value_counts())

elif chart_type == "Scadenze per mese" and "Prossima Scadenza Noleggio" in filtered.columns:
    tmp = filtered.copy()
    tmp = tmp.dropna(subset=["Prossima Scadenza Noleggio"])
    if not tmp.empty:
        tmp["Mese"] = pd.to_datetime(tmp["Prossima Scadenza Noleggio"]).dt.to_period("M").astype(str)
        st.bar_chart(tmp["Mese"].value_counts().sort_index())
    else:
        st.info("Nessuna scadenza disponibile nei dati filtrati.")

elif chart_type == "Distribuzione contratti aperti" and "Tot. Contratti (aperti)" in filtered.columns:
    st.bar_chart(filtered["Tot. Contratti (aperti)"].fillna(0).astype(int).value_counts().sort_index())

st.caption("Suggerimento: cambia i filtri nella barra laterale per aggiornare tabella e grafici.")
