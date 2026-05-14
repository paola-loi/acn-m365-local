# acn-m365-local

MCP server Python che simula il connettore Microsoft 365 di Claude.ai leggendo
direttamente dai file sincronizzati in OneDrive locale.

**Zero OAuth. Zero IT approval. 100% locale.**

---

## Prerequisiti

- Python 3.10+
- Claude Code CLI (`npm install -g @anthropic-ai/claude-code`)
- OneDrive for Business sincronizzato su `C:\Users\{USERNAME}\OneDrive - Accenture\`

---

## Installazione (3 passi)

```powershell
git clone https://github.com/paola-loi/acn-m365-local.git
cd acn-m365-local
.\install.ps1
```

Lo script:
1. Verifica che Python sia installato
2. Installa le dipendenze (`mcp`, `python-docx`, `openpyxl`, `python-pptx`)
3. Registra il server MCP in Claude Code
4. Conferma con `claude mcp list`

Dopo l'installazione, **riavvia Claude Code** e verifica con `/mcp`.

---

## Struttura OneDrive attesa

```
C:\Users\{USERNAME}\OneDrive - Accenture\
├── MARS\
│   ├── kickoff.docx
│   ├── budget_2025.xlsx
│   └── roadmap.pptx
├── Unilever\
│   ├── meeting_notes.md
│   └── deliverables.docx
└── ...
```

Il server legge qualsiasi sottocartella come "progetto".

---

## Formati supportati

| Estensione | Libreria |
|-----------|---------|
| `.docx` | python-docx |
| `.xlsx` | openpyxl |
| `.pptx` | python-pptx |
| `.md` / `.txt` | stdlib |

---

## Esempi di utilizzo

```
"Elenca tutti i file nella cartella MARS"
"Mostrami solo gli Excel nel progetto Unilever"
"Leggi il file MARS/budget_2025.xlsx"
"Cerca 'action items' in tutti i documenti"
"Cerca 'forecast' solo nel progetto MARS"
"Quali file ho modificato questa settimana?"
"Crea un file di note in MARS/riunione_20260505.md"
"Genera un report summary del progetto MARS"
"Elenca i fogli di MARS/budget_2025.xlsx"
"Leggi il range A1:D50 del foglio Forecast in MARS/budget_2025.xlsx"
"Spiega la formula nella cella C12 di MARS/budget_2025.xlsx"
"Aggiungi una colonna Margin = Revenue - Cost in budget_2025.xlsx"
"Cerca dove compare 'forecast' nelle formule di budget_2025.xlsx"
```

> Le scritture su xlsx **non sovrascrivono** mai l'originale: vengono salvate
> accanto come `<nome>.claude.xlsx`. Apri quel file in Excel per vedere le
> formule calcolate (openpyxl scrive solo l'espressione, non il valore).

---

## Tool disponibili

| Tool | Descrizione |
|------|------------|
| `list_project_files` | Lista file in una cartella, con filtro per estensione |
| `read_document` | Legge il contenuto di un file |
| `search_documents` | Cerca keyword nel contenuto di tutti i file |
| `get_recent_files` | File modificati negli ultimi N giorni |
| `create_document` | Crea un nuovo file `.md` o `.docx` |
| `update_document` | Sovrascrive un file esistente |
| `create_summary_report` | Genera report `.md` aggregato di un progetto |
| `list_sheets` | Elenca fogli di un xlsx con dimensioni |
| `describe_sheet` | Header, tipi colonna inferiti, named ranges di un foglio |
| `read_cells` | Legge un range A1 con valori **e** formule |
| `write_cells` | Applica modifiche di celle salvando su copia `<nome>.claude.xlsx` |
| `find_in_xlsx` | Cerca una stringa in valori e formule di un xlsx |
| `view_xlsx` | Restituisce URL del viewer HTML live (auto-reload su write) |

---

## Viewer live (sperimentale)

All'avvio del server parte un piccolo HTTP server su `http://127.0.0.1:8765`
(stdlib, zero dipendenze extra). Bind solo localhost.

Chiedi a Claude *"apri il viewer di MARS/budget.claude.xlsx"* → ti dà l'URL.
La pagina mostra una tabella per ogni foglio e si ricarica automaticamente:

- quando Claude scrive via `write_cells`
- quando salvi il file da LibreOffice/Excel (polling mtime ~1s)

Limiti:
- è un **viewer**, non un editor — niente click sulle celle.
- le formule mostrano l'ultimo valore *cached* da Excel (`openpyxl` non
  ricalcola). Apri il file in Excel/Calc per vedere il calcolo reale.
- truncate a 1000 righe × 100 colonne per foglio (banner se troncato).
- se la porta 8765 è occupata il server tenta 8766–8775, poi una porta
  random — guarda lo stderr del processo per la porta effettiva.

---

## Troubleshooting

| Errore | Soluzione |
|--------|-----------|
| `OneDrive base path not found` | Verifica che OneDrive sia in esecuzione e sincronizzato |
| `File not found` | Controlla il path e aspetta che la sync di OneDrive completi |
| `Permission denied` | Chiudi il file in Office e attendi la sincronizzazione |
| `Module not found` | Riesegui `pip install -r requirements.txt` |
| Server non appare in `/mcp` | Esegui `claude mcp list` e riavvia Claude Code |

---

## Sviluppo locale

```powershell
# Testa l'avvio del server (attende JSON-RPC su stdin — Ctrl+C per uscire)
python server.py

# Oppure tramite CLI MCP
mcp run server.py
```
