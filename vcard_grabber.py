import math
import os
import time
import urllib.parse
import random
import requests
import shutil
import csv
from bs4 import BeautifulSoup

# Basis-URL der Website
BASE_URL = 'https://search.ch'
# URL für die Telefon-Suche (Suchergebnisseite)
SEARCH_URL = BASE_URL + '/tel/'

def human_sleep(base_delay):
    """
    Führt eine Pause mit einem zufälligen Jitter durch,
    um die Anfragen menschlicher wirken zu lassen.
    
    Parameter:
      - base_delay: Basiswartezeit in Sekunden.
    """
    actual_delay = base_delay * random.uniform(0.8, 1.2)
    time.sleep(actual_delay)

def get_user_input():
    """
    Fragt den Benutzer nach den Suchparametern:
      - Sector (misc)
      - Kanton (kanton)
    """
    misc = input("Bitte gib den Sector (misc) ein (z.B. metallbau): ").strip()
    kanton = input("Bitte gib den Kanton (kanton) ein (z.B. AG): ").strip()
    return misc, kanton

def get_delay_settings():
    """
    Fragt den Benutzer nach den gewünschten Wartezeiten:
      - Wartezeit zwischen Detailseiten (in Sekunden)
      - Wartezeit zwischen Suchseiten (in Sekunden)
      - Anzahl der vCards, nach denen eine zusätzliche Pause eingelegt werden soll
      - Dauer der zusätzlichen Pause (in Sekunden)
    Bei ungültiger Eingabe werden Standardwerte verwendet.
    """
    try:
        detail_page_delay = float(input("Bitte gib die Wartezeit (in Sekunden) zwischen Detailseiten an (Standard 1): ") or "1")
        search_page_delay = float(input("Bitte gib die Wartezeit (in Sekunden) zwischen Suchseiten an (Standard 2): ") or "2")
        additional_threshold = int(input("Bitte gib an, nach wie vielen heruntergeladenen vCards eine zusätzliche Pause eingelegt werden soll (Standard 10): ") or "10")
        additional_pause = float(input("Bitte gib die Dauer der zusätzlichen Pause in Sekunden an (Standard 300 = 5 Minuten): ") or "300")
    except ValueError:
        print("Ungültige Eingabe. Es werden Standardwerte verwendet.")
        detail_page_delay = 1
        search_page_delay = 2
        additional_threshold = 10
        additional_pause = 300
    return detail_page_delay, search_page_delay, additional_threshold, additional_pause

def get_search_page(page_number, misc, kanton):
    """
    Lädt die Suchergebnisseite für eine bestimmte Seite und Suchparameter.
    Mit 'firma=1' wird sichergestellt, dass nur Firmenkunden angezeigt werden.
    
    Parameter:
      - page_number: Nummer der Suchergebnisseite
      - misc: Suchbegriff für den Sector
      - kanton: Suchbegriff für den Kanton
    
    Rückgabe:
      Den HTML-Text der Suchergebnisseite.
    """
    params = {
        'misc': misc,
        'kanton': kanton,
        'firma': '1',
        'pages': page_number
    }
    print(f"\nAbrufe Suchseite {page_number} (misc='{misc}', kanton='{kanton}', firma=1)...")
    response = requests.get(SEARCH_URL, params=params)
    response.raise_for_status()
    return response.text

def extract_total_results(html):
    """
    Liest die Gesamtzahl der Ergebnisse aus dem HTML der Suchergebnisseite,
    indem der Inhalt des <span class="tel-result-count"> ausgelesen wird.
    
    Parameter:
      - html: HTML-Text der Suchergebnisseite
    
    Rückgabe:
      Die Anzahl der gefundenen Ergebnisse als Integer.
    """
    soup = BeautifulSoup(html, 'html.parser')
    span = soup.find('span', class_='tel-result-count')
    if span:
        try:
            return int(span.get_text(strip=True))
        except ValueError:
            return 0
    return 0

def extract_detail_urls(html):
    """
    Extrahiert die URLs der Detailseiten aus der Suchergebnisseite.
    Es werden alle <a>-Tags ausgewählt, deren href mit '/tel/' beginnt,
    nicht 'vcard' enthält und nicht mit '/tel/?' startet.
    
    Parameter:
      - html: HTML-Text der Suchergebnisseite
    
    Rückgabe:
      Eine Liste von Detailseiten-URLs.
    """
    soup = BeautifulSoup(html, 'html.parser')
    detail_urls = set()
    for a in soup.find_all('a', href=True):
        href = a['href']
        if href.startswith('/tel/') and not href.startswith('/tel/?') and 'vcard' not in href:
            detail_urls.add(href)
    return list(detail_urls)

def extract_vcard_url(html):
    """
    Sucht auf einer Detailseite nach dem Link zum vCard-Download.
    Es wird ein <a>-Tag gesucht, dessen href '/tel/vcard/' enthält.
    
    Parameter:
      - html: HTML-Text der Detailseite
    
    Rückgabe:
      Die URL zum vCard-Download, falls gefunden, ansonsten None.
    """
    soup = BeautifulSoup(html, 'html.parser')
    vcard_link = soup.find('a', href=lambda href: href and '/tel/vcard/' in href)
    if vcard_link:
        return vcard_link['href']
    return None

def download_vcard(vcard_url, filepath):
    """
    Lädt die vCard von der angegebenen URL herunter und speichert sie unter dem
    angegebenen Dateipfad.
    
    Parameter:
      - vcard_url: URL, unter der die vCard verfügbar ist
      - filepath: Vollständiger Dateipfad, unter dem die vCard gespeichert werden soll
    """
    if vcard_url.startswith('/'):
        vcard_url = BASE_URL + vcard_url
    print(f"vCard herunterladen von {vcard_url} ...")
    response = requests.get(vcard_url)
    response.raise_for_status()
    with open(filepath, 'wb') as f:
        f.write(response.content)

def move_vcards_without_email(output_dir, links_txt_path):
    """
    Überprüft alle vCard-Dateien im angegebenen Ordner und verschiebt jene,
    die keine E-Mail-Adresse enthalten, in einen Unterordner "keine_email".
    Danach wird das Master-Link-File (links.txt) aktualisiert:
      - Einträge zu vCards ohne E-Mail werden in links_keine_email.txt geschrieben.
    
    Parameter:
      - output_dir: Der Ordner, in dem die vCards abgelegt sind.
      - links_txt_path: Pfad zur Master-Link-Datei.
    """
    keine_email_dir = os.path.join(output_dir, "keine_email")
    os.makedirs(keine_email_dir, exist_ok=True)
    
    # Lese die Originaleinträge aus links.txt
    if os.path.exists(links_txt_path):
        with open(links_txt_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    else:
        lines = []
    main_lines = []
    keine_email_lines = []
    
    # Durchsuche alle vCard-Dateien im Hauptordner
    for filename in os.listdir(output_dir):
        if filename.lower().endswith(".vcf"):
            filepath = os.path.join(output_dir, filename)
            try:
                with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                if "email" not in content.lower():
                    new_path = os.path.join(keine_email_dir, filename)
                    shutil.move(filepath, new_path)
                    print(f"VCard '{filename}' enthält keine E-Mail und wurde nach '{keine_email_dir}' verschoben.")
            except Exception as e:
                print(f"Fehler beim Überprüfen von {filename}: {e}")
    
    # Aktualisiere die Link-Dateien
    for line in lines:
        line = line.strip()
        if not line:
            continue
        parts = line.split("|", 1)
        if len(parts) != 2:
            continue
        filename, detail_url = parts
        main_filepath = os.path.join(output_dir, filename)
        keine_email_filepath = os.path.join(keine_email_dir, filename)
        if os.path.exists(main_filepath):
            main_lines.append(f"{filename}|{detail_url}\n")
        elif os.path.exists(keine_email_filepath):
            keine_email_lines.append(f"{filename}|{detail_url}\n")
    
    with open(links_txt_path, 'w', encoding='utf-8') as f:
        f.writelines(main_lines)
    links_keine_email_path = os.path.join(output_dir, "links_keine_email.txt")
    with open(links_keine_email_path, 'w', encoding='utf-8') as f:
        f.writelines(keine_email_lines)
    print(f"links.txt und links_keine_email.txt wurden aktualisiert.")

def parse_vcard(filepath):
    """
    Liest den Inhalt einer vCard (.vcf) und extrahiert relevante Felder.
    Folgende Felder werden übernommen:
      - FN: Full Name
      - ORG: Organisation (Firma)
      - ADR: Adresse (alle Teilfelder werden mit Komma getrennt)
      - TEL: Alle Telefonnummern (durch Komma getrennt)
      - EMAIL: Alle E-Mail-Adressen (durch Komma getrennt)
      - ROLE: Rolle (z.B. Tätigkeit)
      - URL: Homepage oder weitere URL
    Sonderzeichen (z.B. ä, ö, ü) werden über UTF-8 verarbeitet.
    
    Rückgabe:
      Ein Dictionary mit den Schlüsseln:
        "Filename", "Full Name", "Organization", "Address",
        "Telephone", "Email", "Role", "URL"
    """
    data = {
        "Filename": os.path.basename(filepath),
        "Full Name": "",
        "Organization": "",
        "Address": "",
        "Telephone": "",
        "Email": "",
        "Role": "",
        "URL": ""
    }
    try:
        with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
            for line in f:
                line = line.strip()
                if line.upper().startswith("FN:"):
                    data["Full Name"] = line.split(":", 1)[1].strip()
                elif line.upper().startswith("ORG:"):
                    data["Organization"] = line.split(":", 1)[1].strip()
                elif line.upper().startswith("ADR"):
                    parts = line.split(":", 1)
                    if len(parts) == 2:
                        adr = parts[1].strip()
                        adr_parts = adr.split(";")
                        # Zusammenfügen der nicht-leeren Felder
                        data["Address"] = ", ".join([p for p in adr_parts if p])
                elif line.upper().startswith("TEL"):
                    parts = line.split(":", 1)
                    if len(parts) == 2:
                        tel = parts[1].strip()
                        if data["Telephone"]:
                            data["Telephone"] += ", " + tel
                        else:
                            data["Telephone"] = tel
                elif line.upper().startswith("EMAIL"):
                    parts = line.split(":", 1)
                    if len(parts) == 2:
                        email = parts[1].strip()
                        if data["Email"]:
                            data["Email"] += ", " + email
                        else:
                            data["Email"] = email
                elif line.upper().startswith("ROLE:"):
                    parts = line.split(":", 1)
                    if len(parts) == 2:
                        data["Role"] = parts[1].strip()
                elif line.upper().startswith("URL:"):
                    parts = line.split(":", 1)
                    if len(parts) == 2:
                        data["URL"] = parts[1].strip()
    except Exception as e:
        print(f"Fehler beim Parsen der vCard {filepath}: {e}")
    return data

def generate_csv_from_vcards(folder, csv_path):
    """
    Generiert eine CSV-Datei aus allen vCard-Dateien in einem Ordner.
    Die CSV enthält die Spalten:
      Filename, Full Name, Organization, Address, Telephone, Email, Role, URL
    Das Encoding wird auf 'utf-8-sig' gesetzt, damit Umlaute korrekt dargestellt werden.
    """
    entries = []
    for filename in os.listdir(folder):
        if filename.lower().endswith(".vcf"):
            filepath = os.path.join(folder, filename)
            data = parse_vcard(filepath)
            entries.append(data)
    fieldnames = ["Filename", "Full Name", "Organization", "Address", "Telephone", "Email", "Role", "URL"]
    try:
        with open(csv_path, 'w', encoding='utf-8-sig', newline='') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for entry in entries:
                writer.writerow(entry)
        print(f"CSV-Datei erstellt: {csv_path}")
    except Exception as e:
        print(f"Fehler beim Erstellen der CSV-Datei {csv_path}: {e}")

    # Suchparameter abfragen
    misc, kanton = get_user_input()
    
    # Wartezeiten abfragen
    detail_page_delay, search_page_delay, additional_threshold, additional_pause = get_delay_settings()
    
    # Erstelle den Ausgabeordner: "vcards" als oberster Ordner, dann Sector und Kanton
    output_dir = os.path.join("vcards", misc.lower(), kanton.lower())
    os.makedirs(output_dir, exist_ok=True)
    
    # Pfad zur zentralen Master-Link-Datei (Format: "Dateiname|Detailseiten-Link")
    links_txt_path = os.path.join(output_dir, "links.txt")
    
    # Abrufen der ersten Suchseite, um die Gesamtanzahl der Ergebnisse zu ermitteln
    try:
        first_page_html = get_search_page(1, misc, kanton)
    except Exception as e:
        print(f"Fehler beim Abrufen der ersten Suchseite: {e}")
        return

    total_results = extract_total_results(first_page_html)
    if total_results == 0:
        print("Keine Ergebnisse gefunden.")
        return

    detail_urls_first_page = extract_detail_urls(first_page_html)
    if not detail_urls_first_page:
        print("Keine Detailseiten auf der ersten Suchseite gefunden.")
        return

    results_per_page = len(detail_urls_first_page)
    total_pages = math.ceil(total_results / results_per_page)
    print(f"Gesamte Ergebnisse: {total_results}, Ergebnisse pro Seite: {results_per_page}, insgesamt {total_pages} Seiten.")

    downloaded_count = 0
    processed_details = set()

    # Iteriere über alle Suchseiten
    for page in range(1, total_pages + 1):
        print(f"\n--- Seite {page} von {total_pages} ---")
        try:
            page_html = get_search_page(page, misc, kanton)
        except Exception as e:
            print(f"Fehler beim Abrufen der Suchseite {page}: {e}")
            continue

        detail_urls = extract_detail_urls(page_html)
        print(f"Auf Suchseite {page} wurden {len(detail_urls)} Detailseiten gefunden.")

        for detail in detail_urls:
            detail_url = BASE_URL + detail
            if detail_url in processed_details:
                print(f"Detailseite {detail_url} wurde bereits verarbeitet. Überspringe.")
                continue
            processed_details.add(detail_url)
            
            try:
                print(f"Detailseite abrufen: {detail_url}")
                detail_resp = requests.get(detail_url)
                detail_resp.raise_for_status()
                detail_html = detail_resp.text
            except Exception as e:
                print(f"Fehler beim Abrufen der Detailseite {detail_url}: {e}")
                continue

            vcard_link = extract_vcard_url(detail_html)
            if not vcard_link:
                print(f"Kein vCard-Link auf Detailseite gefunden: {detail_url}")
                continue

            parsed = urllib.parse.urlparse(vcard_link)
            filename = os.path.basename(parsed.path)
            vcard_filepath = os.path.join(output_dir, filename)
            keine_email_path = os.path.join(output_dir, "keine_email", filename)

            if os.path.exists(vcard_filepath) or os.path.exists(keine_email_path):
                print(f"{filename} existiert bereits (im Hauptordner oder in 'keine_email'). Überspringe diesen Eintrag.")
            else:
                try:
                    download_vcard(vcard_link, vcard_filepath)
                    downloaded_count += 1
                    print(f"Heruntergeladen: {filename}")
                    with open(links_txt_path, 'a', encoding='utf-8') as txt_file:
                        txt_file.write(f"{filename}|{detail_url}\n")
                except Exception as e:
                    print(f"Fehler beim Herunterladen der vCard von {vcard_link}: {e}")
                    continue

            human_sleep(detail_page_delay)
            if downloaded_count > 0 and downloaded_count % additional_threshold == 0:
                print(f"\nEs wurden {downloaded_count} vCards heruntergeladen. Warte zusätzlich {additional_pause} Sekunden...")
                human_sleep(additional_pause)
        
        human_sleep(search_page_delay)

    print(f"\nFertig! Insgesamt wurden {downloaded_count} vCards heruntergeladen und in '{output_dir}' abgelegt.")
    
    # Verschiebe vCards ohne E-Mail in den Unterordner "keine_email" und aktualisiere die Link-Dateien
    move_vcards_without_email(output_dir, links_txt_path)
    
    # Generiere CSV-Dateien aus den vCard-Inhalten (Adressbuch)
    csv_main = os.path.join(output_dir, "vcards.csv")
    generate_csv_from_vcards(output_dir, csv_main)
    
    keine_email_folder = os.path.join(output_dir, "keine_email")
    csv_keine = os.path.join(output_dir, "vcards_keine_email.csv")
    if os.path.exists(keine_email_folder):
        generate_csv_from_vcards(keine_email_folder, csv_keine)

if __name__ == '__main__':
    main()
