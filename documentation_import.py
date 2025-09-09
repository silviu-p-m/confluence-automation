import os
import requests
import re
from unidecode import unidecode
from pathlib import Path
import traceback


# --- KONFIGURATION ---
USERNAME = "Default"
API_TOKEN = "Default"
CONFLUENCE_URL = "Default"
BASE_DIR = r"C:\Users\Default"
DOCUMENTATION_PAGE = "Dokumentation Stand 05.2025"


FIRMEN_LISTE = [
    "3KV", "Abraxa", "AIB Holding (Hans Sieber GmbH)"
]

# --- SESSION SETUP ---
session = requests.Session()
session.auth = (USERNAME, API_TOKEN)
session.headers.update({
    "Accept": "application/json",
    "Content-Type": "application/json"
})


# --- HILFSFUNKTIONEN ---
def generate_space_key(space_name):
    cleaned = unidecode(space_name).upper()
    cleaned = re.sub(r"[^A-Z0-9]+", "", cleaned)
    return f"DOCS{cleaned[:8]}" if cleaned else f"AUTO{hash(space_name)[:5]}"


def get_page_id(space_key, page_title):
    cql = f'space="{space_key}" and title="{page_title}"'
    try:
        response = session.get(
            f"{CONFLUENCE_URL}/content/search",
            params={"cql": cql}
        )
        results = response.json().get("results", [])
        if results:
            return results[0]["id"]
        # Wenn nicht gefunden, erweitere die Suche
        cql = f'space="{space_key}" and title ~ "{page_title}"'
        response = session.get(
            f"{CONFLUENCE_URL}/content/search",
            params={"cql": cql}
        )
        results = response.json().get("results", [])
        return results[0]["id"] if results else None
    except Exception as e:
        print(f" Fehler bei der Suche nach Seite: {str(e)}")
        return None


def get_or_create_documentation_page(space_key):
    page_id = get_page_id(space_key, DOCUMENTATION_PAGE)
    if page_id:
        return page_id

    # Wenn nicht gefunden, erstelle die Seite
    print(f"️ Dokumentationsseite nicht gefunden, erstelle neu in {space_key}")

    # Finde die Homepage des Spaces
    try:
        response = session.get(
            f"{CONFLUENCE_URL}/space/{space_key}?expand=homepage"
        )
        space_info = response.json()
        home_page_id = space_info.get("homepage", {}).get("id")
        if not home_page_id:
            print(f" Homepage nicht gefunden in {space_key}")
            print(f" Space-Info: {space_info}")
            return None
    except Exception as e:
        print(f" Fehler beim Laden der Homepage: {str(e)}")
        return None

    # Erstelle die Dokumentationsseite
    page_data = {
        "type": "page",
        "title": DOCUMENTATION_PAGE,
        "space": {"key": space_key},
        "ancestors": [{"id": home_page_id}],
        "body": {
            "storage": {
                "value": "<p>Automatisch erstellte Dokumentationsseite</p>",
                "representation": "storage"
            }
        }
    }
    try:
        response = session.post(
            f"{CONFLUENCE_URL}/content",
            json=page_data
        )
        if response.status_code == 200:
            print(f" Dokumentationsseite in {space_key} erstellt")
            return response.json().get("id")
        else:
            print(f" Fehler beim Erstellen der Dokumentationsseite: HTTP {response.status_code}")
            print(f" Response: {response.text}")
    except Exception as e:
        print(f" Fehler beim Erstellen der Dokumentationsseite: {str(e)}")

    return None


def create_page(space_key, parent_id, title, content):
    page_data = {
        "type": "page",
        "title": title,
        "space": {"key": space_key},
        "ancestors": [{"id": parent_id}] if parent_id else [],
        "body": {
            "storage": {
                "value": content,
                "representation": "storage"
            }
        }
    }
    try:
        response = session.post(
            f"{CONFLUENCE_URL}/content",
            json=page_data
        )
        if response.status_code == 200:
            return response.json().get("id")
        else:
            print(f" Fehler beim Erstellen der Seite: HTTP {response.status_code}")
            print(f" Response: {response.text}")
    except Exception as e:
        print(f" Fehler beim Erstellen der Seite: {str(e)}")
    return None


def upload_attachment(page_id, file_path):
    filename = os.path.basename(file_path)
    headers = {"X-Atlassian-Token": "no-check"}

    try:
        with open(file_path, 'rb') as file:
            response = session.post(
                f"{CONFLUENCE_URL}/content/{page_id}/child/attachment",
                headers=headers,
                files={"file": (filename, file)}
            )

        if response.status_code == 200:
            results = response.json().get("results", [])
            if results:
                return f'<p><a href="{results[0]["_links"]["download"]}">Originaldatei herunterladen</a></p>'
    except Exception as e:
        print(f" Fehler beim Hochladen des Anhangs {filename}: {str(e)}")
    return ""


def extract_pdf_content(file_path):
    try:
        print(f" Versuche PDF-Extraktion: {file_path.name}")
        text_content = ""

        with open(file_path, 'rb') as pdf_file:
            # Überprüfe, ob die PDF verschlüsselt ist
            reader = PyPDF2.PdfReader(pdf_file)

            if reader.is_encrypted:
                print(" PDF ist verschlüsselt - versuche leeres Passwort")
                try:
                    reader.decrypt("")
                except:
                    print(" Entschlüsselung fehlgeschlagen")
                    return '<p>Verschlüsselte PDF - Inhalt kann nicht extrahiert werden</p>'

            # Überprüfe die Anzahl der Seiten
            num_pages = len(reader.pages)
            print(f" PDF hat {num_pages} Seiten")

            # Extrahiere Text von jeder Seite
            for i, page in enumerate(reader.pages, 1):
                try:
                    page_text = page.extract_text()
                    if page_text:
                        text_content += f"--- Seite {i} ---\n{page_text}\n\n"
                    else:
                        text_content += f"--- Seite {i} (kein Text) ---\n\n"
                except Exception as e:
                    print(f"️ Fehler bei Seite {i}: {str(e)}")
                    text_content += f"--- Seite {i} (Fehler beim Extrahieren) ---\n\n"

            if not text_content.strip():
                return '<p>PDF enthält keinen extrahierbaren Text (evtl. gescanntes Dokument)</p>'

            return f'<pre>{text_content[:15000]}</pre>'

    except Exception as e:
        print(f" PDF-Extraktion fehlgeschlagen: {str(e)}")
        traceback.print_exc()
        return '<p>PDF-Inhalt konnte nicht extrahiert werden</p>'


def extract_docx_content(file_path):
    try:
        print(f" Verarbeite DOCX: {file_path.name}")
        full_text = []
        doc = Document(file_path)
        for para in doc.paragraphs:
            full_text.append(para.text)
        text_content = "\n".join(full_text)
        return f'<pre>{text_content[:15000]}</pre>'
    except Exception as e:
        print(f" DOCX-Extraktion fehlgeschlagen: {str(e)}")
        return '<p>Word-Inhalt konnte nicht extrahiert werden</p>'


def extract_excel_content(file_path):
    try:
        print(f" Verarbeite Excel: {file_path.name}")
        html_content = ""

        # Verwende openpyxl als Engine für bessere Kompatibilität
        excel_file = pd.ExcelFile(file_path, engine='openpyxl')

        for sheet_name in excel_file.sheet_names:
            df = excel_file.parse(sheet_name, engine='openpyxl')
            html_content += f'<h3>{sheet_name}</h3>'
            html_content += df.head(100).to_html(index=False)  # Nur erste 100 Zeilen

        return html_content
    except Exception as e:
        print(f" Excel-Extraktion fehlgeschlagen: {str(e)}")
        traceback.print_exc()
        return '<p>Excel-Inhalt konnte nicht extrahiert werden</p>'


def generate_file_content(file_path):
    filename = os.path.basename(file_path)
    extension = file_path.suffix.lower()

    content = f'<h2>{filename}</h2>'

    # PDF-Inhalt extrahieren
    if extension == '.pdf':
        content += extract_pdf_content(file_path)

    # Word-Dokument konvertieren
    elif extension == '.docx':
        content += extract_docx_content(file_path)

    # Excel-Inhalt anzeigen
    elif extension in ['.xlsx', '.xls']:
        content += extract_excel_content(file_path)

    return content


def find_company_directory(company_name):
    base_path = Path(BASE_DIR)

    # 1. Versuch: Exakter Match
    exact_match = base_path / company_name
    if exact_match.exists() and exact_match.is_dir():
        return exact_match

    # 2. Versuch: Case-insensitive Suche
    for dir_entry in base_path.iterdir():
        if dir_entry.is_dir() and company_name.lower() == dir_entry.name.lower():
            return dir_entry

    # 3. Versuch: Teilstring-Suche
    for dir_entry in base_path.iterdir():
        if dir_entry.is_dir() and company_name.lower() in dir_entry.name.lower():
            return dir_entry

    return None


# --- HAUPTFUNKTION ---
def main():
    print(" Starte Dokumentenimport...")

    for firma in FIRMEN_LISTE:
        print(f"\n{'━' * 40}")
        print(f" Verarbeite Firma: {firma}")

        # Space-Key generieren
        space_key = generate_space_key(firma)
        print(f" Space-Key: {space_key}")

        # Dokumentationsseite finden oder erstellen
        doc_page_id = get_or_create_documentation_page(space_key)
        if not doc_page_id:
            print(f" Dokumentationsseite konnte nicht erstellt werden in {space_key}")
            continue

        # Firmenordner finden
        company_dir = find_company_directory(firma)
        if not company_dir:
            print(f"️ Ordner für {firma} nicht gefunden")
            continue

        print(f" Ordner: {company_dir}")

        # Dokumente rekursiv suchen
        file_count = 0
        for root, _, files in os.walk(company_dir):
            for file in files:
                file_path = Path(root) / file
                extension = file_path.suffix.lower()

                if extension in ['.pdf', '.docx', '.xlsx', '.xls']:
                    file_count += 1
                    print(f"\n Verarbeite Datei ({file_count}): {file_path.name}")

                    try:
                        # Inhalte generieren
                        file_content = generate_file_content(file_path)

                        # Anhang hochladen
                        download_link = upload_attachment(doc_page_id, file_path)
                        full_content = file_content + (download_link if download_link else '')

                        # Seite erstellen
                        page_title = f"Dokument: {file_path.name}"
                        page_id = create_page(
                            space_key=space_key,
                            parent_id=doc_page_id,
                            title=page_title,
                            content=full_content
                        )

                        if page_id:
                            print(f" Seite erstellt: {page_title}")
                        else:
                            print(f" Fehler beim Erstellen der Seite für {file_path.name}")

                    except Exception as e:
                        print(f" Kritischer Fehler bei {file_path.name}: {str(e)}")
                        traceback.print_exc()

        print(f" Gesamte Dateien verarbeitet: {file_count}")

    print("\n Import abgeschlossen!")


if __name__ == "__main__":
    # Überprüfen der Abhängigkeiten
    try:
        import PyPDF2
        import pandas as pd
        from docx import Document

        print(" Alle benötigten Bibliotheken sind installiert")
        main()
    except ImportError as e:
        print(f" Fehlende Bibliothek: {str(e)}")
        print("Bitte installieren Sie die fehlenden Pakete mit:")
        print("pip install PyPDF2 pandas python-docx openpyxl")