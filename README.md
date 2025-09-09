# Confluence Automation

Python-Skripte zur Arbeit mit der Confluence REST API:  
Spaces klonen, Firmenordner importieren (PDF/DOCX/XLSX) und auto-generierte Dokumentationsseiten anlegen.

## Stack
Python · requests · PyPDF2 · python-docx · pandas · openpyxl · unidecode

## Run
```bash
pip install -r requirements.txt
cp .env.example .env   # User, Token, Base URL eintragen
python CopySpaceConfluence.py --source ENG --target ENG-CLONE
