---
name: acn-m365-local
description: |
  Accesses Microsoft 365 files (Word, Excel, PowerPoint, Markdown) synced
  locally from OneDrive - Accenture. Use this skill when the user asks to
  read project documents, search files, or create reports/summaries from
  local OneDrive folders. No authentication required: direct local
  filesystem access.
---

## Typical workflows

### Project exploration

```
"List all files in the MARS folder"
→ list_project_files(project="MARS")

"Show me only the Excel files in the Unilever project"
→ list_project_files(project="Unilever", file_type=".xlsx")

"Which files are in Unilever/2025/Q1?"
→ list_project_files(project="Unilever/2025/Q1")
```

### Reading documents

```
"Read the file MARS/budget_Q1.xlsx"
→ read_document(path="MARS/budget_Q1.xlsx")

"Show me the notes from the March 15 meeting"
→ search_documents(query="March 15") then read_document(...)

"Open the presentation MARS/kickoff.pptx"
→ read_document(path="MARS/kickoff.pptx")
```

### Cross-file search

```
"Find 'action items' across all my documents"
→ search_documents(query="action items")

"Search for 'forecast' only in the MARS project"
→ search_documents(query="forecast", project="MARS")

"Find every document that mentions 'Rossi'"
→ search_documents(query="Rossi")
```

### Recent files

```
"Which files did I modify this week?"
→ get_recent_files(days=7)

"Show me changes from the last 3 days in the Unilever project"
→ get_recent_files(days=3, project="Unilever")

"Files updated today"
→ get_recent_files(days=1)
```

### Create and update

```
"Create a notes file for today's meeting at MARS/meeting_20260505.md"
→ create_document(path="MARS/meeting_20260505.md", content="...", doc_format="md")

"Create a Word document with the minutes at Unilever/minutes.docx"
→ create_document(path="Unilever/minutes.docx", content="...", doc_format="docx")

"Update MARS/todo.md with these items"
→ update_document(path="MARS/todo.md", content="...")
```

### Summary reports

```
"Generate a summary report of the MARS project"
→ create_summary_report(project="MARS", output_path="MARS/summary_report.md")

"Create a summary of all Unilever documents"
→ create_summary_report(project="Unilever", output_path="Unilever/_summary.md")
```
