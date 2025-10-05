import streamlit as st
import pandas as pd
from datetime import date, timedelta

st.set_page_config(page_title="Gestione Clienti", layout="wide")

st.title("📒 Gestione Clienti — Dashboard & Schede")

st.markdown("""
Questa applicazione ti permette di **gestire i clienti**, analizzare le **scadenze**
e visualizzare le **schede dettagliate** direttamente dal file Excel `GESTIONE_CLIENTI.xlsm`.
""")

# ---------------------------
# FUNZIONE DI CARICAMENTO DATI
# ---------------------------
def load_from_excel(file):
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
    df = df.dropna(how="all")

    # --- parsing robusto delle date ---
    def _parse_date(x):
        if pd.isna(x):
            return ""
        try:
            return pd.to_datetime(str(x), dayfirst=True, errors="coerce").strftime("%d/%m/%Y")
        except Exception:
            return ""
    for c in ["Ultimo Recall", "Ultima Visita", "Prossima Scadenza Noleggio"]:
        if c in df.columns:
            df[c] = df[c].apply(_parse_date)

    if "Tot. Contratti (aperti)" in df.columns:
        df["Tot. Contratti (aperti)"] = pd.to_numeric(df["Tot. Contratti (aperti)"], errors="coerce")

    if "Cliente" in df.columns:
        df = df[df["Cliente"].notna()]

    return df


# ---------------------------
# UPLOAD FILE
# ---------------------------
uploaded = st.file_uploader("📂 Carica il file Excel (.xlsm / .xlsx)", type=["xlsm", "xlsx"])

if not uploaded:
    st.info("⬆️ Carica il file `GESTIONE_CLIENTI.xlsm` per iniziare.")
    st.stop()

try:
    df = load_from_excel(uploaded)
except Exception as e:
    st.error(f"Errore nel caricamento del file: {e}")
    st.stop()

# ---------------------------
# TAB PRINCIPALI
# ---------------------------
tab1, tab2 = st.tabs(["📊 Dashboard", "📇 Schede Cliente"])

# ---------------------------
# TAB 1 — DASHBOARD
# ---------------------------
with tab1:
    st.sidebar.header("🔎 Filtri")

    query = st.sidebar.text_input("Cerca cliente (nome parziale)", "")
    tmk_values = sorted([x for x in df["TMK"].dropna().unique()]) if "TMK" in df.columns else []
    tmk_sel = st.sidebar.multiselect("TMK", tmk_values, default=tmk_values)
    days = st.sidebar.slider("Mostra scadenze entro (giorni)", 0, 365, 60, 15)

    today = date.today()
    deadline_limit = today + timedelta(days=days)

    filtered = df.copy()

    if query:
        filtered = filtered[filtered["Cliente"].astype(str).str.contains(query, case=False, na=False)]

    if "TMK" in filtered.columns and tmk_sel:
        filtered = filtered[filtered["TMK"].isin(tmk_sel)]

    if "Prossima Scadenza Noleggio" in filtered.columns:
        temp = pd.to_datetime(filtered["Prossima Scadenza Noleggio"], errors="coerce", dayfirst=True)
        mask = temp.isna() | (temp <= pd.to_datetime(deadline_limit))
        filtered = filtered[mask]

    # KPI
    k1, k2, k3 = st.columns(3)
    with k1:
        st.metric("Clienti mostrati", len(filtered))
    with k2:
        tot_aperti = filtered["Tot. Contratti (aperti)"].sum() if "Tot. Contratti (aperti)" in filtered.columns else 0
        st.metric("Totale contratti aperti", int(tot_aperti))
    with k3:
        if "Prossima Scadenza Noleggio" in filtered.columns:
            count_scadenze = pd.to_datetime(filtered["Prossima Scadenza Noleggio"], errors="coerce", dayfirst=True)
            count_scadenze = count_scadenze[count_scadenze <= pd.to_datetime(deadline_limit)].count()
            st.metric(f"Scadenze entro {days} giorni", int(count_scadenze))

    st.subheader("📋 Elenco Clienti (dopo i filtri)")
    st.dataframe(filtered, use_container_width=True)

    csv = filtered.to_csv(index=False).encode("utf-8")
    st.download_button("💾 Scarica come CSV", csv, "clienti_filtrati.csv", "text/csv")

    # GRAFICI
    st.subheader("📈 Grafici")
    chart_type = st.selectbox("Tipo di grafico", ["Clienti per TMK", "Scadenze per mese", "Distribuzione contratti"])

    if chart_type == "Clienti per TMK" and "TMK" in filtered.columns:
        st.bar_chart(filtered["TMK"].value_counts())

    elif chart_type == "Scadenze per mese" and "Prossima Scadenza Noleggio" in filtered.columns:
        tmp = filtered.copy()
        tmp = tmp.dropna(subset=["Prossima Scadenza Noleggio"])
        if not tmp.empty:
            tmp["Mese"] = pd.to_datetime(tmp["Prossima Scadenza Noleggio"], dayfirst=True, errors="coerce").dt.to_period("M").astype(str)
            st.bar_chart(tmp["Mese"].value_counts().sort_index())
        else:
            st.info("Nessuna scadenza disponibile nei dati filtrati.")
    elif chart_type == "Distribuzione contratti" and "Tot. Contratti (aperti)" in filtered.columns:
        st.bar_chart(filtered["Tot. Contratti (aperti)"].fillna(0).astype(int).value_counts().sort_index())

# ---------------------------
# TAB 2 — SCHEDE CLIENTE
# ---------------------------
with tab2:
    st.subheader("📇 Schede Cliente Dettagliate")

    cliente_sel = st.selectbox("Seleziona un cliente:", df["Cliente"].unique())

    if cliente_sel:
        try:
            sheet = pd.read_excel(uploaded, sheet_name=str(cliente_sel), header=None, engine="openpyxl").fillna("")

            # Rimuovi righe iniziali inutili
            if sheet.iloc[0].astype(str).str.contains("Torna all'Indice", case=False).any():
                sheet = sheet.iloc[1:].reset_index(drop=True)

            st.markdown(f"### 📘 Scheda Cliente — **{cliente_sel}**")

            # Dati cliente
            st.markdown("#### 🧾 Dati Cliente")
            for i, row in sheet.head(15).iterrows():
                label = str(row[0]).strip()
                value = str(row[1]).strip() if len(row) > 1 else ""
                if label and value and label.lower() != "nan":
                    st.markdown(f"**{label}:** {value}")

            # Contratti
            idx_contratti = sheet.index[
                sheet.astype(str).apply(lambda r: r.str.contains("Contratti di Noleggio", case=False).any(), axis=1)
            ]
            if not idx_contratti.empty:
                start_row = idx_contratti[0] + 1
                contratti_raw = sheet.iloc[start_row:].reset_index(drop=True)
                header_idx = contratti_raw.index[
                    contratti_raw.astype(str).apply(lambda r: r.str.contains("DATA INIZIO", case=False).any(), axis=1)
                ]
                if not header_idx.empty:
                    contratti_raw = contratti_raw.iloc[header_idx[0]:].reset_index(drop=True)

                contratti = contratti_raw.copy()
                contratti.columns = contratti.iloc[0]
                contratti = contratti[1:].dropna(how="all").reset_index(drop=True)
                contratti = contratti.loc[:, ~contratti.columns.duplicated()]

                # formatta le date
                for col in contratti.columns:
                    if any(k in str(col).lower() for k in ["data", "inizio", "fine"]):
                        contratti[col] = pd.to_datetime(contratti[col], errors="coerce", dayfirst=True).dt.strftime("%d/%m/%Y")

                st.markdown("#### 📑 Contratti di Noleggio")
                st.dataframe(contratti, use_container_width=True)

            # Note cliente
            idx_note = sheet.index[
                sheet.astype(str).apply(lambda r: r.str.contains("NOTE CLIENTI", case=False).any(), axis=1)
            ]
            if not idx_note.empty:
                note_text = ""
                for i in range(idx_note[0] + 1, len(sheet)):
                    line = " ".join(str(x) for x in sheet.iloc[i] if pd.notna(x)).strip()
                    if line:
                        note_text += line + "\n"
                if note_text.strip():
                    st.markdown("#### 🗒️ Note Cliente")
                    st.text_area("Note", note_text, height=150)

        except Exception as e:
            st.error(f"❌ Non riesco a caricare la scheda per {cliente_sel}. Errore: {e}")
