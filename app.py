import re
import io
import json
import unicodedata
from typing import Tuple, Optional, Dict
import pandas as pd
import streamlit as st

# ======================= Look & feel "app" =======================
st.set_page_config(
    page_title="Gestione Clienti",
    page_icon="icon-512.png",    # se la tieni in /static usa "static/icon-512.png"
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

st.title("📄 Gestione Clienti — Contratti & Note")

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

# =========== Parsing foglio cliente: contratti + note ===========
def parse_contracts_and_notes(sheet_df: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
    df = sheet_df.copy()
    df = df.dropna(axis=1, how="all")
    df = df.astype(str).where(~df.isna(), None)

    col0 = df.columns[0]
    first_col = df[col0].apply(lambda x: str(x).strip() if x is not None else "")

    # "Contratti di Noleggio" -> header sulla riga successiva
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
            # stop su NOTE CLIENTI
            if normalize_text(row0).startswith("note clienti"):
                break
            # stop su riga vuota
            if all((str(x).strip() == "" or str(x).strip().lower() == "none") for x in df.iloc[r].tolist()):
                break
            rows.append([None if str(x).strip().lower() == "none" else x for x in df.iloc[r].tolist()])
        if rows:
            contratti_df = pd.DataFrame(rows, columns=headers)
            # parse base per date (colonne che contengono "data")
            for c in list(contratti_df.columns):
                if "data" in normalize_text(c):
                    contratti_df[c] = pd.to_datetime(contratti_df[c], errors="coerce", dayfirst=True)

    # NOTE CLIENTI (riga dopo l’intestazione)
    note_text = ""
    for idx, key in first_col.items():
        if normalize_text(key).startswith("note clienti"):
            rr = idx + 1
            if rr < len(df):
                note_text = get_first_nonempty([df.at[rr, c] for c in df.columns])
            break

    return contratti_df, note_text

# ---------- INFO CLIENTE (chiavi/valori prima dei contratti) ----------
def parse_client_info(sheet_df: pd.DataFrame) -> Tuple[str, Dict[str, str]]:
    df = sheet_df.copy()
    df = df.dropna(axis=1, how="all")
    df = df.astype(str).where(~df.isna(), None)

    col0 = df.columns[0]
    first_col = df[col0].apply(lambda x: str(x).strip() if x is not None else "")

    # riga di stop: contratti o note
    stop_rows = []
    for idx, key in first_col.items():
        nk = normalize_text(key)
        if nk.startswith("contratti di noleggio") or nk.startswith("note clienti"):
            stop_rows.append(idx)
    stop_at = min(stop_rows) if stop_rows else len(df)

    # nome cliente
    nome = ""
    for idx, key in first_col.items():
        if idx >= stop_at:
            break
        nk = normalize_text(key)
        if nk in ("nome cliente", "cliente"):
            nome = get_first_nonempty([df.at[idx, c] for c in df.columns[1:]])
            if nome:
                break

    # info chiave->valore
    info: Dict[str, str] = {}
    SKIP = {"scheda cliente", "torna all indice", "totale contratti", "dati cliente", "cliente", "nome cliente"}
    for idx, key in first_col.items():
        if idx >= stop_at:
            break
        k_raw = str(key).strip()
        if not k_raw:
            continue
        if normalize_text(k_raw) in SKIP:
            continue
        v = get_first_nonempty([df.at[idx, c] for c in df.columns[1:]])
        if v:
            info[k_raw] = v

    return nome, info

# =============== Indice: elenco clienti & match =================
def extract_client_list_from_indice(indice_df: pd.DataFrame) -> list:
    if indice_df.empty:
        return []
    header_row0 = indice_df.iloc[0].to_dict()
    candidates = [c for c, v in header_row0.items() if isinstance(v, str) and "cliente" in v.lower()]
    if not candidates:
        candidates = [c for c in indice_df.columns if indice_df[c].notna().any()]
        if not candidates:
            return []
    col = candidates[0]
    values = (
        indice_df[col]
        .iloc[1:]
        .dropna()
        .astype(str)
        .map(str.strip)
        .tolist()
    )
    values = [v for v in values if normalize_text(v) not in ("", "cliente")]
    seen, out = set(), []
    for v in values:
        if v not in seen:
            seen.add(v)
            out.append(v)
    return out

def find_client_sheet_name(sheets: Dict[str, pd.DataFrame], cliente: str) -> Optional[str]:
    target = normalize_text(cliente)
    for name in sheets.keys():
        if normalize_text(name) == target:
            return name
    for name in sheets.keys():
        if normalize_text(name).startswith(target):
            return name
    for name in sheets.keys():
        if target in normalize_text(name):
            return name
    return None

# =============== Export XLSX cliente (foglio singolo) ===============
def export_client_xlsx(cliente: str, info: Dict[str, str], contratti: pd.DataFrame, note: str) -> bytes:
    """
    Genera un XLSX con:
      - sezione 'Dati Cliente' (chiave/valore)
      - sezione 'Contratti di Noleggio'
      - sezione 'NOTE CLIENTI'
    NB: le macro .xlsm non sono mantenute (si crea sempre un .xlsx).
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # 1) Info (2 colonne)
        info_df = pd.DataFrame(list(info.items()), columns=["Campo", "Valore"])
        info_df.to_excel(writer, sheet_name=cliente[:31] or "Cliente", index=False, startrow=0)
        ws = writer.sheets[cliente[:31] or "Cliente"]

        # 2) Titolo contratti
        start_row = len(info_df) + 2
        ws.cell(row=start_row, column=1, value="Contratti di Noleggio")
        # 3) Tabella contratti
        if not contratti.empty:
            contr_out = contratti.copy()
            # format date -> stringa per Excel
            for c in contr_out.columns:
                if pd.api.types.is_datetime64_any_dtype(contr_out[c]):
                    contr_out[c] = contr_out[c].dt.strftime("%d/%m/%Y")
            contr_out.to_excel(writer, sheet_name=ws.title, index=False, startrow=start_row)
            end_row = start_row + len(contr_out) + 2
        else:
            end_row = start_row + 2

        # 4) Note
        ws.cell(row=end_row, column=1, value="NOTE CLIENTI")
        ws.cell(row=end_row + 1, column=1, value=note or "")

    buf.seek(0)
    return buf.read()

# ============================ APP FLOW ============================
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

# Trova "Indice"
names_map = {n.lower(): n for n in sheets_dict.keys()}
indice_key = names_map.get("indice")
clienti = extract_client_list_from_indice(sheets_dict[indice_key]) if indice_key else []

# Ricerca + selezione cliente
col1, col2 = st.columns([1, 2])
with col1:
    query = st.text_input("🔎 Cerca cliente (da 'Indice')", "", placeholder="digita per filtrare…")
def _match(n: str, q: str) -> bool:
    return normalize_text(q) in normalize_text(n) if q else True
filtered_clienti = [c for c in clienti if _match(c, query)]
with col2:
    st.caption(f"{len(filtered_clienti)} clienti trovati" if clienti else "Nessun cliente in 'Indice'")

if filtered_clienti:
    cliente_sel = st.selectbox("Seleziona cliente", ["-- Seleziona --"] + filtered_clienti, index=0)
else:
    st.warning("Nessun cliente disponibile (controlla il foglio 'Indice').")
    cliente_sel = None

# Persistenza semplice delle modifiche in sessione
if "notes_store" not in st.session_state:
    st.session_state.notes_store = {}  # {cliente -> nota}
if "info_store" not in st.session_state:
    st.session_state.info_store = {}   # {cliente -> dict info}
if "contracts_store" not in st.session_state:
    st.session_state.contracts_store = {}  # {cliente -> DataFrame}

# ------------------------ Visualizzazione scheda ------------------------
if cliente_sel and cliente_sel != "-- Seleziona --":
    foglio = find_client_sheet_name(sheets_dict, cliente_sel)
    if not foglio:
        st.warning("Foglio cliente non trovato.")
        st.stop()

    sheet_df = sheets_dict[foglio]

    # Parse dati
    nome_cli, info_cli = parse_client_info(sheet_df)
    contratti_df, note_esistente = parse_contracts_and_notes(sheet_df)

    # Carica eventuali modifiche precedenti dalla sessione
    info_cli = st.session_state.info_store.get(cliente_sel, info_cli)
    contratti_df = st.session_state.contracts_store.get(cliente_sel, contratti_df)
    note_val = st.session_state.notes_store.get(cliente_sel, note_esistente or "")

    st.markdown(f"### 🧾 {cliente_sel}")
    if nome_cli and normalize_text(nome_cli) != normalize_text(cliente_sel):
        st.caption(f"Nome da scheda: {nome_cli}")

    # -------- Toggle MODIFICA --------
    edit_mode = st.toggle("✏️ Modifica/Inserisci dati", value=False, help="Abilita la maschera per modificare dati, contratti e note")

    # ====== MODALITÀ LETTURA ======
    if not edit_mode:
        if info_cli:
            st.subheader("👤 Dati Cliente")
            ordered = ["Indirizzo", "Città", "CAP", "TELEFONO", "MAIL", "RIF.", "RIF 2.", "IBAN", "partita iva", "SDI", "Ultimo Recall", "ultima visita"]
            keys = [k for k in ordered if k in info_cli] + [k for k in info_cli.keys() if k not in ordered]
            c1, c2 = st.columns(2)
            for i, k in enumerate(keys):
                target = c1 if i % 2 == 0 else c2
                with target:
                    st.markdown(f"**{k}**")
                    st.write(info_cli[k])

        st.subheader("📑 Contratti di Noleggio")
        if not contratti_df.empty:
            display_df = contratti_df.copy()
            for c in display_df.columns:
                if pd.api.types.is_datetime64_any_dtype(display_df[c]):
                    display_df[c] = display_df[c].dt.strftime("%d/%m/%y")
            st.dataframe(display_df, use_container_width=True)
        else:
            st.info("Nessun contratto trovato in questa scheda.")

        st.subheader("📝 Note Cliente")
        st.write(note_val or "—")

    # ====== MODALITÀ MODIFICA (MASCHERA) ======
    else:
        st.subheader("👤 Dati Cliente — modifica")
        # campi suggeriti + quelli extra trovati
        base_order = ["Indirizzo", "Città", "CAP", "TELEFONO", "MAIL", "RIF.", "RIF 2.", "IBAN", "partita iva", "SDI", "Ultimo Recall", "ultima visita"]
        extra_keys = [k for k in info_cli.keys() if k not in base_order]
        keys = base_order + extra_keys

        # form
        with st.form("edit_info"):
            c1, c2 = st.columns(2)
            new_info = {}
            for i, k in enumerate(keys):
                target = c1 if i % 2 == 0 else c2
                with target:
                    new_info[k] = st.text_input(k, value=info_cli.get(k, ""))
            submitted_info = st.form_submit_button("💾 Salva Dati Cliente (sessione)")
        if submitted_info:
            st.session_state.info_store[cliente_sel] = new_info
            info_cli = new_info
            st.success("Dati cliente aggiornati nella sessione.")

        st.subheader("📑 Contratti di Noleggio — modifica")
        # abilita aggiunta/eliminazione righe
        editable = contratti_df.copy()
        # Converte eventuali datetime in stringa editabile
        for c in editable.columns:
            if pd.api.types.is_datetime64_any_dtype(editable[c]):
                editable[c] = editable[c].dt.strftime("%d/%m/%Y")
        edited = st.data_editor(
            editable,
            num_rows="dynamic",
            use_container_width=True,
            key=f"editor_{cliente_sel}"
        )
        if st.button("💾 Salva Contratti (sessione)"):
            # Riprova a parse delle date
            out = edited.copy()
            for c in out.columns:
                if "data" in normalize_text(c):
                    out[c] = pd.to_datetime(out[c], errors="coerce", dayfirst=True)
            st.session_state.contracts_store[cliente_sel] = out
            contratti_df = out
            st.success("Contratti aggiornati nella sessione.")

        st.subheader("📝 Note Cliente — modifica")
        new_note = st.text_area("Testo note", value=note_val, height=140, placeholder="Scrivi o aggiorna le note qui…")
        if st.button("💾 Salva Nota (sessione)"):
            st.session_state.notes_store[cliente_sel] = new_note
            note_val = new_note
            st.success("Nota aggiornata nella sessione.")

        # ---- Esportazioni ----
        st.divider()
        cA, cB, cC = st.columns(3)
        with cA:
            csv = contratti_df.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Scarica Contratti (CSV)", data=csv, file_name=f"contratti_{normalize_text(cliente_sel)}.csv", mime="text/csv")
        with cB:
            payload = {"cliente": cliente_sel, "info": info_cli, "note": note_val}
            st.download_button("⬇️ Scarica Info+Note (JSON)", data=json.dumps(payload, ensure_ascii=False, indent=2), file_name=f"note_info_{normalize_text(cliente_sel)}.json", mime="application/json")
        with cC:
            xls_bytes = export_client_xlsx(cliente_sel, info_cli, contratti_df, note_val)
            st.download_button("⬇️ Esporta XLSX (foglio cliente)", data=xls_bytes, file_name=f"{normalize_text(cliente_sel)}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.stop()
