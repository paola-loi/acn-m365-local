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
```

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
