---
name: acn-m365-local
description: |
  Accede ai file Microsoft 365 (Word, Excel, PowerPoint, Markdown) sincronizzati
  localmente da OneDrive - Accenture. Usa questo skill quando l'utente chiede di
  leggere documenti di progetto, cercare file, creare report o summary da cartelle
  OneDrive locali. Non richiede autenticazione: accesso diretto al filesystem locale.
---

## Workflow tipici

### Esplorazione progetto

```
"Elenca tutti i file nella cartella MARS"
→ list_project_files(project="MARS")

"Mostrami solo gli Excel nel progetto Unilever"
→ list_project_files(project="Unilever", file_type=".xlsx")

"Quali file ci sono in Unilever/2025/Q1?"
→ list_project_files(project="Unilever/2025/Q1")
```

### Lettura documenti

```
"Leggi il file MARS/budget_Q1.xlsx"
→ read_document(path="MARS/budget_Q1.xlsx")

"Mostrami le note della riunione del 15 marzo"
→ search_documents(query="15 marzo") poi read_document(...)

"Apri la presentazione MARS/kickoff.pptx"
→ read_document(path="MARS/kickoff.pptx")
```

### Ricerca trasversale

```
"Cerca 'action items' in tutti i miei documenti"
→ search_documents(query="action items")

"Cerca 'forecast' solo nel progetto MARS"
→ search_documents(query="forecast", project="MARS")

"Trova tutti i documenti che menzionano 'Rossi'"
→ search_documents(query="Rossi")
```

### File recenti

```
"Quali file ho modificato questa settimana?"
→ get_recent_files(days=7)

"Mostrami le modifiche degli ultimi 3 giorni nel progetto Unilever"
→ get_recent_files(days=3, project="Unilever")

"File aggiornati oggi"
→ get_recent_files(days=1)
```

### Creazione e aggiornamento

```
"Crea un file di note per la riunione di oggi in MARS/riunione_20260505.md"
→ create_document(path="MARS/riunione_20260505.md", content="...", doc_format="md")

"Crea un documento Word con il verbale in Unilever/verbale.docx"
→ create_document(path="Unilever/verbale.docx", content="...", doc_format="docx")

"Aggiorna il file MARS/todo.md con questi punti"
→ update_document(path="MARS/todo.md", content="...")
```

### Report summary

```
"Genera un report summary del progetto MARS"
→ create_summary_report(project="MARS", output_path="MARS/summary_report.md")

"Crea un riepilogo di tutti i documenti Unilever"
→ create_summary_report(project="Unilever", output_path="Unilever/_summary.md")
```
