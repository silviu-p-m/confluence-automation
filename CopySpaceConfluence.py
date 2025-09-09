import requests
import re
from unidecode import unidecode

# Link für den API Key https://id.atlassian.com/manage-profile/security/api-tokens

# --- KONFIGURATION ---
SOURCE_SPACE_KEY = "BVL"
USERNAME = "test"
API_TOKEN = "default"
EXCLUDED_PAGES = {"Bereichsvorlage Home"}
FIRMEN_LISTE = [
    "3KV", "Abraxa", "AIB Holding (Hans Sieber GmbH)"
]

CONFLUENCE_URL = "https://microcat-service.atlassian.net/wiki/rest/api"

# --- INITIALISIERUNG ---
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


def space_exists(space_key):
    try:
        response = session.get(f"{CONFLUENCE_URL}/space/{space_key}")
        return response.status_code == 200
    except requests.exceptions.HTTPError:
        return False


def create_space(space_key, space_name):
    space_data = {
        "key": space_key,
        "name": space_name,
        "description": {
            "plain": {
                "value": "Automatisch generierter Bereich",
                "representation": "plain"
            }
        }
    }
    try:
        response = session.post(
            f"{CONFLUENCE_URL}/space",
            json=space_data,
            headers={"X-Atlassian-Token": "no-check"}
        )
        response.raise_for_status()
        print(f" Space {space_key} erstellt")
        return True
    except Exception as e:
        print(f" FEHLER: {str(e)}")
        if hasattr(e, 'response'):
            print(f" Response-Details: {e.response.text}")
        return False


def get_all_space_pages(space_key):
    start = 0
    limit = 50
    all_pages = []

    while True:
        response = session.get(
            f"{CONFLUENCE_URL}/content/search",
            params={
                "cql": f"space={space_key} and type=page order by created",
                "expand": "body.storage,ancestors,version",
                "start": start,
                "limit": limit
            }
        )
        data = response.json()
        all_pages.extend([
            p for p in data.get("results", [])
            if p["title"] not in EXCLUDED_PAGES
        ])

        if data.get("size", 0) < limit:
            break
        start += limit

    # Sortierung nach Hierarchieebene
    return sorted(all_pages, key=lambda x: len(x.get("ancestors", [])))


def create_page(page_data):
    try:
        response = session.post(
            f"{CONFLUENCE_URL}/content",
            json=page_data
        )
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f" Seitenfehler: {str(e)}")
        return None


def copy_space_content(source_key, target_key):
    pages = get_all_space_pages(source_key)
    id_mapping = {}

    print(f"\n Starte Kopiervorgang für {target_key}")
    print(f" Gefundene Seiten: {len(pages)}")

    for page in pages:
        try:
            # Body-Content extrahieren
            body_content = page.get("body", {}).get("storage", {}).get("value", "")

            # Neue Seitenstruktur
            new_page = {
                "type": "page",
                "title": page["title"],
                "space": {"key": target_key},
                "body": {
                    "storage": {
                        "value": body_content,
                        "representation": "storage"
                    }
                }
            }

            # Hierarchie beibehalten
            if ancestors := page.get("ancestors"):
                parent_id = ancestors[-1]["id"]
                if parent_id in id_mapping:
                    new_page["ancestors"] = [{"id": id_mapping[parent_id]}]

            # Seite erstellen
            created = create_page(new_page)

            if created and "id" in created:
                id_mapping[page["id"]] = created["id"]
                print(f" {page['title']}")
            else:
                print(f" Fehler bei: {page['title']}")

        except Exception as e:
            print(f"️ Kritischer Fehler: {str(e)}")

    print(f" Erfolgreich kopiert: {len(pages)} Seiten nach {target_key}\n")


def main():
    print(" Starte vollständigen Kopiervorgang...\n")

    for firma in FIRMEN_LISTE:
        target_key = generate_space_key(firma)
        print(f"\n{'━' * 40}")
        print(f" Firma: {firma}")
        print(f" Space-Key: {target_key}")

        if space_exists(target_key):
            print(f" Space existiert bereits: {target_key}")
            continue

        if create_space(target_key, firma):
            copy_space_content(SOURCE_SPACE_KEY, target_key)


if __name__ == "__main__":
    main()
    print("\n Prozess abgeschlossen! Überprüfe die Bereiche in Confluence.")