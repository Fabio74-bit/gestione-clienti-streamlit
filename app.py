import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 🧭 CONFIGURAZIONE BASE
st.set_page_config(page_title="Gestione Clienti — Dashboard completa", layout="wide")

# 🧾 TITOLO E DESCRIZIONE
st.title("📒 Gestione Clienti — Dashboard completa")

st.markdown("""
Questa app mostra i dati dei clienti caricati manualmente da un file Excel 📂  
Carica il file aggiornato ogni volta che vuoi per visualizzare i dati più recenti.
""")

# 📤 UPLOAD FILE
uploaded_file = st.file_uploader("Seleziona il file Excel (.xlsm o .xlsx)", type=["xlsm", "xlsx"])

if uploaded_file is not None:
    try:
        # Legge il file Excel in un DataFrame pandas
        df = pd.read_excel(uploaded_file)

        # ✅ Conferma caricamento
        st.success(f"✅ File caricato con successo alle {datetime.now().strftime('%H:%M:%S')}!")

        # 📊 Mostra anteprima dati
        st.subheader("Anteprima dei dati caricati")
        st.dataframe(df, use_container_width=True)

        # 🔍 Filtri opzionali (esempio base)
        with st.expander("Filtra i dati"):
            colonne = df.columns.tolist()
            colonna_scelta = st.selectbox("Scegli una colonna da filtrare:", colonne)
            valore = st.text_input("Inserisci un valore da cercare:")
            if valore:
                risultati = df[df[colonna_scelta].astype(str).str.contains(valore, case=False, na=False)]
                st.write(f"🔎 **{len(risultati)} risultati trovati**")
                st.dataframe(risultati, use_container_width=True)

        # 📈 Statistiche base (opzionale)
        with st.expander("Statistiche generali"):
            st.write("Numero totale di righe:", len(df))
            st.write("Numero di colonne:", len(df.columns))

    except Exception as e:
        st.error(f"❌ Errore nel caricamento del file: {e}")

else:
    st.warning("📄 Nessun file caricato. Carica un file Excel per iniziare.")
