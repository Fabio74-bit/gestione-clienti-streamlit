import re
import json
import unicodedata
from typing import Tuple, Optional, Dict, List
import pandas as pd
import streamlit as st

# ======================= Look & feel "app" =======================
st.set_page_config(
    page_title="Gestione Clienti",
    page_icon="icon-512.png",      # cambia in "static/icon-512.png" se la tieni in /static
    layout="wide"
)
st.markdown("""
<link rel="apple-touch-icon" sizes="180x180" href="/apple-touch-icon.png">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="Gestione Clienti">
<style>
#MainMenu, header, footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# =========================== Utility ============================
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

def clean_table(df: pd.DataFrame) -> pd.DataFrame:
    """Rimuove colonne/righe completamente vuote o 'None', sostituisce 'None' con ''."""
    if df is None or df.empty:
        return pd.DataFrame()
    tmp = df.copy()
    tmp = tmp.astype(object).where(~tmp.isna(), None)
    tmp = tmp.applymap(lambda x: "" if (x is None or str(x).strip().lower() == "none") else x)
    # drop colonne vuote
    tmp = tmp.replace("", pd.NA).dropna(axis=1, how="all").fillna("")
    # drop righe vuote
    tmp = tmp.replace("", pd.NA).dropna(axis=0, how="all").fillna("")
    return tmp

# ======= Parser: Contratti di Noleggio + Note (da foglio cliente) =======
def parse_contracts_and_notes(sheet_df: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
    df = sheet_df.copy()
    df = df.dropna(axis=1, how="all")
    df = df.astype(str).where(~df.isna(), None)

    col0 = df.columns[0]
    first_col = df[col0].apply(lambda x: str(x).strip() if x is not None else "")

    # trova intestazione "Contratti di Noleggio"
    header_row = None
    for idx, key in first_col.items():
        if normalize_text(key).startswith("contratti di noleggio"):
            header_row = idx + 1
            break

    contratti_df = pd.DataFrame()
    if header_row is not None and header_row < len(df):
        headers = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[header_row].tolist()]
        headers = [h if h else f"Col_{i}" for i, h in enumerate(headers)]
        rows = []
        for r in range(header_row + 1, len(df)):
            row0 = str(df.iloc[r, 0]).strip() if pd.notna(df.iloc[r, 0]) else ""
            if normalize_text(row0).startswith("note clienti"):
                break
            if all((str(x).strip() == "" or str(x).strip().lower() == "none") for x in df.iloc[r].tolist()):
                continue
            rows.append([None if str(x).strip().lower() == "none" else x for x in df.iloc[r].tolist()])
        if rows:
            contratti_df = pd.DataFrame(rows, columns=headers)
            # parse date
            for c in list(contratti_df.columns):
                if "data" in normalize_text(c):
                    contratti_df[c] = pd.to_datetime(contratti_df[c], errors="coerce", dayfirst=True)

    # NOTE CLIENTI (riga successiva al titolo)
    note_text = ""
    for idx, key in first_col.items():
        if normalize_text(key).startswith("note clienti"):
            rr = idx + 1
            if rr < len(df):
                note_text = get_first_nonempty([df.at[rr, c] for c in df.columns])
            break

    # pulizia contratti
    if not contratti_df.empty:
        for c in contratti_df.columns:
            if pd.api.types.is_datetime64_any_dtype(contratti_df[c]):
                contratti_df[c] = contratti_df[c].dt.strftime("%d/%m/%y")
        contratti_df = clean_table(contratti_df)

    return contratti_df, note_text

# ---------- Parser: Dati cliente (chiave/valore prima dei contratti) ----------
def parse_client_info(sheet_df: pd.DataFrame) -> Tuple[str, Dict[str, str]]:
    df = sheet_df.copy()
    df = df.dropna(axis=1, how="all")
    df = df.astype(str).where(~df.isna(), None)

    col0 = df.columns[0]
    first_col = df[col0].apply(lambda x: str(x).strip() if x is not None else "")

    # righe di stop
    stop_rows = []
    for idx, key in first_col.items():
        nk = normalize_text(key)
        if nk.startswith("contratti di noleggio") or nk.startswith("note clienti"):
            stop_rows.append(idx)
    stop_at = min(stop_rows) if stop_rows else len(df)

    # nome cliente
    nome = ""
    for idx, key in first_col.items():
        if idx >= stop_at: break
        nk = normalize_text(key)
        if nk in ("nome cliente", "cliente"):
            nome = get_first_nonempty([df.at[idx, c] for c in df.columns[1:]])
            if nome: break

    # chiave -> valore
    info: Dict[str, str] = {}
    SKIP = {"scheda cliente", "torna all indice", "totale contratti", "dati cliente", "cliente", "nome cliente"}
    for idx, key in first_col.items():
        if idx >= stop_at: break
        k_raw = str(key).strip()
        if not k_raw: continue
        if normalize_text(k_raw) in SKIP: continue
        v = get_first_nonempty([df.at[idx, c] for c in df.columns[1:]])
        if v:
            vv = "" if str(v).strip().lower() == "none" else str(v).strip()
            if vv:
                info[k_raw] = vv

    return nome, info

# ===== Indice: elenco dall'Indice + filtri per nome e città =====
def extract_client_list_from_indice(indice_df: pd.DataFrame) -> pd.DataFrame:
    """Ritorna DataFrame con colonne: Nome, Città (se presente), Telefono (se presente)."""
    if indice_df is None or indice_df.empty:
        return pd.DataFrame(columns=["Nome", "Città", "Telefono"])

    header_row0 = indice_df.iloc[0].to_dict()
    col_cli = None
    for c, v in header_row0.items():
        if isinstance(v, str) and "cliente" in v.lower():
            col_cli = c
            break
    if col_cli is None:
        candidates = [c for c in indice_df.columns if indice_df[c].notna().any()]
        col_cli = candidates[0] if candidates else indice_df.columns[0]

    col_citta = None
    col_tel = None
    for c in indice_df.columns:
        v0 = str(indice_df.at[0, c]) if 0 in indice_df.index else ""
        nk = normalize_text(v0)
        if not col_citta and ("citta" in nk or "citt" in nk): col_citta = c
        if not col_tel and ("telefono" in nk or "tel" in nk): col_tel = c

    names = indice_df[col_cli].iloc[1:].dropna().astype(str).map(str.strip).tolist() if col_cli in indice_df.columns else []
    cities = indice_df[col_citta].iloc[1:].astype(str).map(str.strip).tolist() if col_citta and col_citta in indice_df.columns else []
    tels = indice_df[col_tel].iloc[1:].astype(str).map(str.strip).tolist() if col_tel and col_tel in indice_df.columns else []

    maxlen = max(len(names), len(cities), len(tels), 0)
    def safe(lst, i): return lst[i] if i < len(lst) else ""
    rows, seen = [], set()
    for i in range(maxlen):
        nome = safe(names, i)
        if not nome or normalize_text(nome) in ("", "cliente"): continue
        key = normalize_text(nome)
        if key in seen: continue
        seen.add(key)
        rows.append({"Nome": nome, "Città": safe(cities, i), "Telefono": safe(tels, i)})
    out = pd.DataFrame(rows, columns=["Nome", "Città", "Telefono"]).fillna("")
    return out

def find_client_sheet_name(sheets: Dict[str, pd.DataFrame], cliente: str) -> Optional[str]:
    target = normalize_text(cliente)
    for name in sheets.keys():
        if normalize_text(name) == target: return name
    for name in sheets.keys():
        if normalize_text(name).startswith(target): return name
    for name in sheets.keys():
        if target in normalize_text(name): return name
    return None

# ============================ STATE & NAV ============================
if "view" not in st.session_state:
    st.session_state.view = "index"   # "index" | "detail"
if "selected_cliente" not in st.session_state:
    st.session_state.selected_cliente = None

def go_index():
    st.session_state.view = "index"
    st.session_state.selected_cliente = None

def go_detail(nome: str):
    st.session_state.view = "detail"
    st.session_state.selected_cliente = nome

# ============================ APP FLOW ============================
uploaded = st.file_uploader("📥 Carica il file Excel (.xlsx/.xlsm)", type=["xlsx", "xlsm"])
if not uploaded:
    st.title("📄 Gestione Clienti — Indice")
    st.info("Carica il file per iniziare.")
    st.stop()

@st.cache_data(show_spinner=False)
def load_all_sheets(file) -> Dict[str, pd.DataFrame]:
    return pd.read_excel(file, sheet_name=None, engine="openpyxl", dtype=str)

sheets_dict = load_all_sheets(uploaded)
if not sheets_dict:
    st.error("Nessun foglio trovato nel file.")
    st.stop()

# Trova "Indice"
names_map = {n.lower(): n for n in sheets_dict.keys()}
indice_key = names_map.get("indice")
indice_df = sheets_dict[indice_key] if indice_key else pd.DataFrame()
index_table = extract_client_list_from_indice(indice_df)

# Persistenza note in sessione
if "notes_store" not in st.session_state:
    st.session_state.notes_store = {}  # {cliente -> nota}

# =============== VIEW: INDEX ===============
if st.session_state.view == "index":
    st.title("📇 Indice Clienti")

    colf1, colf2, colf3 = st.columns([2,2,1])
    with colf1:
        q_name = st.text_input("🔎 Cerca per nome", value="", placeholder="es. Rossi, 2 ESSE…")
    with colf2:
        q_city = st.text_input("🏙️ Cerca per città", value="", placeholder="es. Milano, Casarile…")
    with colf3:
        st.write(""); st.write("")
        if st.button("🔄 Pulisci filtri"):
            q_name = ""; q_city = ""
            st.rerun()

    filt = index_table.copy()
    if q_name:
        filt = filt[filt["Nome"].astype(str).map(lambda x: normalize_text(q_name) in normalize_text(x))]
    if q_city:
        filt = filt[filt["Città"].astype(str).map(lambda x: normalize_text(q_city) in normalize_text(x))]

    st.caption(f"{len(filt)} clienti trovati")
    st.dataframe(filt, use_container_width=True, hide_index=True)

    choices: List[str] = ["-- Seleziona --"] + filt["Nome"].tolist()
    pick = st.selectbox("Apri scheda cliente", choices)
    if pick and pick != "-- Seleziona --":
        go_detail(pick)
        st.rerun()

# =============== VIEW: DETAIL ===============
elif st.session_state.view == "detail":
    cliente_sel: Optional[str] = st.session_state.selected_cliente
    st.button("← Torna all’Indice", on_click=go_index)

    if not cliente_sel:
        st.warning("Nessun cliente selezionato.")
        st.stop()

    foglio = find_client_sheet_name(sheets_dict, cliente_sel)
    if not foglio:
        st.warning("Foglio cliente non trovato.")
        st.stop()

    sheet_df = sheets_dict[foglio]
    nome_cli, info_cli = parse_client_info(sheet_df)
    contratti_df, note_esistente = parse_contracts_and_notes(sheet_df)
    note_val = st.session_state.notes_store.get(cliente_sel, note_esistente or "")

    st.markdown(f"## 🧾 {cliente_sel}")
    if nome_cli and normalize_text(nome_cli) != normalize_text(cliente_sel):
        st.caption(f"Nome da scheda: {nome_cli}")

    # Dati Cliente SOPRA
    if info_cli:
        st.subheader("👤 Dati Cliente")
        ordered = ["Indirizzo", "Città", "CAP", "TELEFONO", "MAIL", "RIF.", "RIF 2.", "IBAN", "partita iva", "SDI", "Ultimo Recall", "ultima visita"]
        keys = [k for k in ordered if k in info_cli] + [k for k in info_cli.keys() if k not in ordered]
        for k in keys:
            st.markdown(f"**{k}**")
            st.write(info_cli[k])
    else:
        st.caption("Nessun dato anagrafico trovato.")

    # Contratti SOTTO (larghi e puliti)
    st.subheader("📑 Contratti di Noleggio")
    if contratti_df is not None and not contratti_df.empty:
        pretty = clean_table(contratti_df)
        st.dataframe(pretty, use_container_width=True, hide_index=True)
    else:
        st.info("Nessun contratto trovato in questa scheda.")

    # Note
    st.subheader("📝 Note Cliente")
    new_note = st.text_area("Testo note", value=note_val, height=140, placeholder="Scrivi o aggiorna le note qui…")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 Salva nota (solo in questa sessione)"):
            st.session_state.notes_store[cliente_sel] = new_note
            st.success("Nota salvata nella sessione corrente.")
    with c2:
        notes_json = json.dumps(st.session_state.notes_store, ensure_ascii=False, indent=2)
        st.download_button("⬇️ Esporta tutte le note (JSON)", data=notes_json, file_name="note_clienti.json", mime="application/json")
