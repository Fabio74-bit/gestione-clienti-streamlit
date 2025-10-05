import re
import json
import unicodedata
from typing import Tuple, Optional, Dict
import pandas as pd
import streamlit as st

# ------------------------------ Utility ------------------------------
def normalize_text(s: str) -> str:
    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("ASCII")
    s = s.lower()
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def get_first_nonempty(values) -> str:
    for v in values:
        if pd.notna(v):
            vs = str(v).strip()
            if vs and vs.lower() != "none":
                return vs
    return ""

# ----------------------- Parsing scheda cliente -----------------------
def parse_contracts_and_notes(sheet_df: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
    """
    Estrae:
      - tabella 'Contratti di Noleggio'
      - 'NOTE CLIENTI' (una riga di testo)
    Torna (contratti_df, note_text).
    """
    df = sheet_df.copy()
    df = df.dropna(axis=1, how="all")
    df = df.astype(str).where(~df.isna(), None)

    # prima colonna considerata come "etichette" della riga
    col0 = df.columns[0]
    first_col = df[col0].apply(lambda x: str(x).strip() if x is not None else "")

    # Trova riga intestazione "Contratti di Noleggio"
    header_row = None
    for idx, key in first_col.items():
        if normalize_text(key).startswith("contratti di noleggio"):
            header_row = idx + 1  # la riga subito sotto contiene gli header
            break

    contratti_df = pd.DataFrame()
    if header_row is not None and header_row < len(df):
        headers = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[header_row].tolist()]
        headers = [h if h else f"Col_{i}" for i, h in enumerate(headers)]

        rows = []
        for r in range(header_row + 1, len(df)):
            row0 = str(df.iloc[r, 0]).strip() if pd.notna(df.iloc[r, 0]) else ""
            # Se incontriamo "NOTE CLIENTI" ci fermiamo
            if normalize_text(row0).startswith("note clienti"):
                break
            # Stop se riga completamente vuota
            if all((str(x).strip() == "" or str(x).strip().lower() == "none") for x in df.iloc[r].tolist()):
                break
            rows.append([None if str(x).strip().lower() == "none" else x for x in df.iloc[r].tolist()])

        if rows:
            contratti_df = pd.DataFrame(rows, columns=headers)
            # parsing base per date: colonne che contengono "data"
            for c in list(contratti_df.columns):
                if "data" in normalize_text(c):
                    contratti_df[c] = pd.to_datetime(contratti_df[c], errors="coerce", dayfirst=True)

    # Trova "NOTE CLIENTI" e prendi la riga successiva (prima cella non vuota)
    note_text = ""
    for idx, key in first_col.items():
        if normalize_text(key).startswith("note clienti"):
            rr = idx + 1
            if rr < len(df):
                note_text = get_first_nonempty([df.at[rr, c] for c in df.columns])
            break

    return contratti_df, note_text

# --------------------------- Indice & match ---------------------------
def extract_client_list_from_indice(indice_df: pd.DataFrame) -> list:
    """Cerca in riga 0 la colonna che contiene 'Cliente' e ritorna i nomi sotto."""
    if indice_df.empty:
        return []
    header_row0 = indice_df.iloc[0].to_dict()
    candidates = [c for c, v in header_row0.items() if isinstance(v, str) and "cliente" in v.lower()]
    if not candidates:
        # fallback: prima colonna non vuota
        candidates = [c for c in indice_df.columns if indice_df[c].notna().any()]
        if not candidates:
            return []
    col = candidates[0]
    values = (
        indice_df[col]
        .iloc[1:]  # salta la riga header
        .dropna()
        .astype(str)
        .map(str.strip)
        .tolist()
    )
    # rimuovi header ripetuti/righe vuote
    values = [v for v in values if normalize_text(v) not in ("", "cliente")]
    seen, out = set(), []
    for v in values:
        if v not in seen:
            seen.add(v)
            out.append(v)
    return out

def find_client_sheet_name(sheets: Dict[str, pd.DataFrame], cliente: str) -> Optional[str]:
    """Trova il foglio del cliente con confronti normalizzati (esatto, inizia con, contiene)."""
    target = normalize_text(cliente)
    # esatto
    for name in sheets.keys():
        if normalize_text(name) == target:
            return name
    # inizia con (gestisce suffissi tipo ' (3)')
    for name in sheets.keys():
        if normalize_text(name).startswith(target):
            return name
    # contiene
    for name in sheets.keys():
        if target in normalize_text(name):
            return name
    return None

# ============================== APP ==============================
st.set_page_config(page_title="Schede Clienti — Contratti & Note", layout="wide")
st.title("📄 Schede Clienti — Contratti & Note")

uploaded = st.file_uploader("📥 Carica il file Excel (.xlsx/.xlsm)", type=["xlsx", "xlsm"])
if not uploaded:
    st.info("Carica il file per iniziare.")
    st.stop()

@st.cache_data(show_spinner=False)
def load_all_sheets(file) -> Dict[str, pd.DataFrame]:
    return pd.read_excel(file, sheet_name=None, engine="openpyxl", dtype=str)

sheets_dict = load_all_sheets(uploaded)
if not sheets_dict:
    st.error("Nessun foglio trovato nel file.")
    st.stop()

# trova "Indice"
names_map = {n.lower(): n for n in sheets_dict.keys()}
indice_key = names_map.get("indice")
clienti = extract_client_list_from_indice(sheets_dict[indice_key]) if indice_key else []

# --- Ricerca cliente ---
col1, col2 = st.columns([1, 2])
with col1:
    query = st.text_input("🔎 Cerca cliente (da 'Indice')", "", placeholder="digita per filtrare…")
def _match(n: str, q: str) -> bool:
    return normalize_text(q) in normalize_text(n) if q else True
filtered_clienti = [c for c in clienti if _match(c, query)]

with col2:
    st.caption(f"{len(filtered_clienti)} clienti trovati" if clienti else "Nessun cliente in 'Indice'")

# --- Selezione cliente ---
if filtered_clienti:
    cliente_sel = st.selectbox("Seleziona cliente", ["-- Seleziona --"] + filtered_clienti, index=0)
else:
    st.warning("Nessun cliente disponibile (controlla il foglio 'Indice').")
    cliente_sel = None

# ------------------------ Gestione NOTE (persistenza semplice) ------------------------
# session_state: dizionario {cliente -> nota}
if "notes_store" not in st.session_state:
    st.session_state.notes_store = {}  # type: ignore

with st.expander("📦 Import/Export note", expanded=False):
    up = st.file_uploader("Carica un JSON di note (facoltativo)", type=["json"], key="upload_notes_json")
    if up is not None:
        try:
            incoming = json.load(up)
            if isinstance(incoming, dict):
                st.session_state.notes_store.update(incoming)
                st.success("Note importate.")
            else:
                st.error("Formato JSON non valido (atteso dict {cliente: nota}).")
        except Exception as e:
            st.error(f"Errore nel parsing del JSON: {e}")
    notes_json = json.dumps(st.session_state.notes_store, ensure_ascii=False, indent=2)
    st.download_button("⬇️ Scarica note (JSON)", data=notes_json, file_name="note_clienti.json", mime="application/json")

# ------------------------ Visualizzazione scheda ------------------------
if cliente_sel and cliente_sel != "-- Seleziona --":
    foglio = find_client_sheet_name(sheets_dict, cliente_sel)
    if not foglio:
        st.warning("Foglio cliente non trovato.")
        st.stop()

    sheet_df = sheets_dict[foglio]
    contratti_df, note_esistente = parse_contracts_and_notes(sheet_df)

    st.markdown(f"### 🧾 {cliente_sel}")

    # ---- Contratti di Noleggio ----
    st.subheader("📑 Contratti di Noleggio")
    if not contratti_df.empty:
        display_df = contratti_df.copy()
        for c in display_df.columns:
            if pd.api.types.is_datetime64_any_dtype(display_df[c]):
                display_df[c] = display_df[c].dt.strftime("%d/%m/%y")
        st.dataframe(display_df, use_container_width=True)
    else:
        st.info("Nessun contratto trovato in questa scheda.")

    # ---- NOTE CLIENTI ----
    st.subheader("📝 Note Cliente")
    current_note = st.session_state.notes_store.get(cliente_sel, note_esistente or "")
    new_note = st.text_area("Testo note", value=current_note, height=140, placeholder="Scrivi o aggiorna le note qui…")

    colA, colB = st.columns([1,1])
    with colA:
        if st.button("💾 Salva nota (solo per questa sessione)"):
            st.session_state.notes_store[cliente_sel] = new_note
            st.success("Nota salvata nella sessione corrente.")
    with colB:
        st.caption("Suggerimento: usa **Scarica note (JSON)** per conservarle e ricaricarle più tardi.")
else:
    st.stop()
