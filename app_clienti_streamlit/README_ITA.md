
# 📒 Dashboard "Gestione Clienti" (super semplice)

Questa app **non richiede esperienza**. Bastano 10 minuti.

## ✅ Cosa fa
- Legge il foglio **Indice** del file `GESTIONE_CLIENTI.xlsm`
- Mostra una tabella filtrabile e ricercabile
- Calcola 3 indicatori (KPI) rapidi
- Disegna grafici semplici (nessuna configurazione complicata)

## ▶️ Come si avvia (passo passo)

1. **Installa Python** (se non c'è già)
   - Windows/Mac: scaricalo da https://www.python.org (scegli la versione 3.10 o superiore).
   - Durante l'installazione, spunta **"Add Python to PATH"** su Windows.

2. **Scarica i file dell'app** e mettili in una cartella insieme al tuo `GESTIONE_CLIENTI.xlsm`.

3. **Apri il Terminale/Prompt dei comandi** in quella cartella.

4. **Installa le librerie** (lo fai una sola volta):
   ```bash
   pip install -r requirements.txt
   ```

5. **Avvia l'app**:
   ```bash
   streamlit run app.py
   ```

6. Si apre una pagina nel browser (se non si apre da sola, vai su `http://localhost:8501`).

## 📂 Cosa c'è dentro
- `app.py` → il codice dell'app
- `requirements.txt` → l'elenco delle librerie da installare
- (facoltativo) `GESTIONE_CLIENTI.xlsm` → metti qui il tuo file, **oppure caricalo** dall'app

## ℹ️ Note utili
- L'app cerca di trovare il foglio **Indice** e di leggere le colonne più importanti:
  - `Cliente`, `Ultimo Recall`, `Ultima Visita`, `Prossima Scadenza Noleggio`, `Tot. Contratti (aperti)`, `TMK`
- Se alcune colonne mancano o hanno nomi diversi, l'app ignora quelle mancanti e mostra le altre disponibili.
- Se cambi il file Excel, **non serve toccare il codice**: ricarica la pagina e basta.

Se vuoi, posso personalizzarla (altri grafici, esportazioni, stampe, ecc.).
