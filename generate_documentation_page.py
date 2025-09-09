import requests
import re
from unidecode import unidecode

# --- KONFIGURATION ---
USERNAME = "Default"
API_TOKEN = "Default"
CONFLUENCE_URL = "Default"
FIRMEN_LISTE = [
    "3KV", "Abraxa", "AIB Holding (Hans Sieber GmbH)"
]

PAGE_TITLE = "Dokumentation Stand 05.2025"

# --- SESSION ---
session = requests.Session()
session.auth = (USERNAME, API_TOKEN)
session.headers.update({
    "Accept": "application/json",
    "Content-Type": "application/json"
})


def generate_space_key(space_name):
    cleaned = unidecode(space_name).upper()
    cleaned = re.sub(r"[^A-Z0-9]+", "", cleaned)
    return f"DOCS{cleaned[:8]}" if cleaned else f"AUTO{hash(space_name)[:5]}"


def get_home_page_id(space_key):
    try:
        response = session.get(
            f"{CONFLUENCE_URL}/space/{space_key}?expand=homepage"
        )
        response.raise_for_status()
        return response.json().get("homepage", {}).get("id")
    except Exception as e:
        print(f" Fehler beim Laden der Home-Seite für {space_key}: {e}")
        return None


def create_documentation_page(space_key, parent_id):
    payload = {
        "type": "page",
        "title": PAGE_TITLE,
        "space": {"key": space_key},
        "ancestors": [{"id": parent_id}],
        "body": {
            "storage": {
                "value": "",
                "representation": "storage"
            }
        }
    }

    try:
        response = session.post(
            f"{CONFLUENCE_URL}/content",
            json=payload
        )
        response.raise_for_status()
        print(f" Dokumentationsseite in {space_key} erstellt")
    except Exception as e:
        print(f" Fehler beim Erstellen der Seite in {space_key}: {e}")


def main():
    print(" Starte das Anfügen leerer Dokumentationsseiten...\n")
    for firma in FIRMEN_LISTE:
        space_key = generate_space_key(firma)
        print(f"\n Bearbeite Firma: {firma} → {space_key}")

        parent_id = get_home_page_id(space_key)
        if parent_id:
            create_documentation_page(space_key, parent_id)
        else:
            print(f"️ Keine gültige Home-Seite für {space_key} gefunden – übersprungen.")

    print("\n Prozess abgeschlossen – alle Seiten wurden verarbeitet.")


if __name__ == "__main__":
    main()
