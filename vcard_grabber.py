import requests
import time
import random
import csv
import os
import json
from datetime import datetime
from bs4 import BeautifulSoup

# --- API configuration ---
API_KEY = "API--KEY--GOES--HERE"  # Your API key
BASE_URL = "https://tel.search.ch/api/"
USAGE_FILE = "api_usage.json"
MONTHLY_QUOTA = 1000

# --- Helper functions ---
def human_sleep(base_delay):
    """Sleep for base_delay seconds with ±20% jitter."""
    actual_delay = base_delay * random.uniform(0.8, 1.2)
    time.sleep(actual_delay)

def load_api_usage():
    """Load API usage info from a JSON file. Reset counter if month has changed."""
    current_month = datetime.now().strftime("%Y-%m")
    try:
        with open(USAGE_FILE, "r") as f:
            usage = json.load(f)
        if usage.get("month") != current_month:
            usage = {"month": current_month, "calls": 0}
    except Exception:
        usage = {"month": current_month, "calls": 0}
    return usage

def save_api_usage(usage):
    """Save API usage info to a JSON file."""
    with open(USAGE_FILE, "w") as f:
        json.dump(usage, f)

def increment_api_counter():
    """Increment the API call counter and save it. Returns the new total."""
    usage = load_api_usage()
    usage["calls"] += 1
    save_api_usage(usage)
    return usage["calls"]

def show_remaining_api_calls():
    """Display how many API calls are remaining for this month."""
    usage = load_api_usage()
    remaining = MONTHLY_QUOTA - usage.get("calls", 0)
    print(f"You have {remaining} API calls remaining for this month.")
    if remaining <= 0:
        print("WARNING: Your API quota has been exhausted!")
    return remaining

def get_user_input():
    """
    Ask the user for search parameters.
    'was' is the general search string (e.g. metallbau).
    'wo' is the geographic refinement (e.g. SO).
    """
    was = input("Bitte gib den Suchbegriff ein (z.B. metallbau): ").strip()
    wo = input("Bitte gib den Ort/Kanton ein (z.B. SO): ").strip()
    return was, wo

def fetch_results(was, wo, pos=1, maxnum=10):
    """
    Fetch a batch of listings from the API.
    - 'was': general search string.
    - 'wo': geographic search.
    - 'pos': starting position (first result is 1).
    - 'maxnum': number of results per call.
    Returns a BeautifulSoup-parsed XML (Atom) feed.
    """
    usage = load_api_usage()
    if usage.get("calls", 0) >= MONTHLY_QUOTA:
        raise Exception("API quota exceeded for this month!")
    
    params = {
        "key": API_KEY,
        "was": was,
        "wo": wo,
        "pos": pos,
        "maxnum": maxnum,
        "lang": "de"
    }
    response = requests.get(BASE_URL, params=params)
    response.raise_for_status()
    increment_api_counter()
    # Parse the Atom XML response using the lxml parser
    soup = BeautifulSoup(response.content, "xml")
    return soup

def sanitize_filename(filename):
    """
    Replace characters that are not allowed in Windows filenames.
    Invalid characters: <>:"/\\|?*
    """
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, "_")
    return filename

def download_vcard(vcard_url, output_folder, entry_id):
    """
    Download a vCard from vcard_url and save it in output_folder.
    Uses entry_id to create a filename if needed.
    """
    if not vcard_url:
        return None
    
    print(f"Downloading vCard from: {vcard_url}")
    resp = requests.get(vcard_url)
    resp.raise_for_status()
    
    # Remove query parameters and get the basename
    filename = os.path.basename(vcard_url.split("?")[0])
    filename = sanitize_filename(filename)
    
    if not filename.lower().endswith(".vcf"):
        filename = f"{entry_id}.vcf"
    
    filepath = os.path.join(output_folder, filename)
    with open(filepath, 'wb') as f:
        f.write(resp.content)
    
    return filepath

def parse_entry(entry):
    """
    Parse an <entry> element from the Atom feed and extract fields.
    Returns a dictionary with the desired fields.
    """
    company = entry.find("tel:org")
    company = company.text if company else ""
    
    firstname = entry.find("tel:firstname")
    firstname = firstname.text if firstname else ""
    
    lastname = entry.find("tel:name")
    lastname = lastname.text if lastname else ""
    
    street = entry.find("tel:street")
    street = street.text if street else ""
    
    streetno = entry.find("tel:streetno")
    streetno = streetno.text if streetno else ""
    
    zip_ = entry.find("tel:zip")
    zip_ = zip_.text if zip_ else ""
    
    city = entry.find("tel:city")
    city = city.text if city else ""
    
    address = f"{street} {streetno}, {zip_} {city}".strip(", ")
    
    phone_elems = entry.find_all("tel:phone")
    phone_numbers = [elem.text.strip() for elem in phone_elems if elem.text]
    phone = ", ".join(phone_numbers)
    
    email = ""
    for extra in entry.find_all("tel:extra"):
        if extra.get("type") and extra.get("type").lower() == "email":
            email = extra.text.strip()
            break
    
    vcard_link = entry.find("link", {"type": "text/x-vcard"})
    vcard_url = vcard_link["href"] if vcard_link and vcard_link.has_attr("href") else ""
    
    entry_id = entry.find("id")
    entry_id = entry_id.text if entry_id else "unknown"
    
    updated = entry.find("updated")
    updated = updated.text if updated else ""
    
    return {
        "EntryId": entry_id,
        "Updated": updated,
        "Company": company,
        "Firstname": firstname,
        "Lastname": lastname,
        "Address": address,
        "Phone": phone,
        "Email": email,
        "VCardUrl": vcard_url,
        "VCardPath": ""  # to be filled when downloaded
    }

def load_existing_entries(csv_path):
    """
    Load existing entries from the master CSV file.
    Returns a dictionary keyed by EntryId.
    """
    entries = {}
    if os.path.exists(csv_path):
        with open(csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                entry_id = row.get("EntryId")
                if entry_id:
                    entries[entry_id] = row
    return entries

def write_csv(csv_path, fieldnames, rows):
    """
    Write rows (a list of dictionaries) to a CSV file with given fieldnames.
    This function filters out any keys not in fieldnames.
    """
    with open(csv_path, 'w', encoding='utf-8-sig', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            filtered_row = {key: row.get(key, "") for key in fieldnames}
            writer.writerow(filtered_row)

# --- Main function ---
def main():
    # Show remaining API calls
    remaining = show_remaining_api_calls()
    if remaining <= 0:
        return
    
    was, wo = get_user_input()
    # Use lowercase for folder names regardless of input
    base_folder = os.path.join(was.lower(), wo.lower())
    csv_folder = os.path.join(base_folder, "csv")
    vcards_email_folder = os.path.join(base_folder, "vcards_email")
    vcards_noemail_folder = os.path.join(base_folder, "vcards_noemail")
    
    # Create folders if they don't exist
    for folder in [base_folder, csv_folder, vcards_email_folder, vcards_noemail_folder]:
        os.makedirs(folder, exist_ok=True)
    
    master_csv = os.path.join(csv_folder, "results_master.csv")
    csv_with_email = os.path.join(csv_folder, "vcards_with_email.csv")
    csv_without_email = os.path.join(csv_folder, "vcards_without_email.csv")
    
    # Load existing entries (if any) to avoid re-scanning
    master_entries = load_existing_entries(master_csv)
    
    pos = 1      # API result position starts at 1
    maxnum = 10  # Number of results per API call
    total_results = None

    while True:
        try:
            soup = fetch_results(was, wo, pos=pos, maxnum=maxnum)
        except Exception as ex:
            print(f"Error during API call: {ex}")
            break
        
        total_tag = soup.find("openSearch:totalResults")
        total_results = int(total_tag.text) if total_tag and total_tag.text.isdigit() else 0
        
        entries = soup.find_all("entry")
        if not entries:
            print("No more entries found.")
            break
        
        print(f"Fetched {len(entries)} entries, total results: {total_results}, starting at pos {pos}")
        
        for entry in entries:
            data = parse_entry(entry)
            eid = data["EntryId"]
            # Check if entry exists in master; update only if the remote 'Updated' is newer.
            if eid in master_entries:
                if data["Updated"] > master_entries[eid].get("Updated", ""):
                    print(f"Updating entry {eid} (newer version found).")
                    if data["VCardUrl"]:
                        folder_to_use = vcards_email_folder if data["Email"].strip() != "" else vcards_noemail_folder
                        try:
                            vcard_path = download_vcard(data["VCardUrl"], folder_to_use, data["EntryId"])
                            data["VCardPath"] = vcard_path
                        except Exception as ex:
                            print(f"vCard download failed for {data['VCardUrl']}: {ex}")
                    master_entries[eid] = data
                else:
                    # If vCard file is missing, try to download it
                    if data["VCardUrl"] and (not master_entries[eid].get("VCardPath") or not os.path.exists(master_entries[eid].get("VCardPath"))):
                        folder_to_use = vcards_email_folder if data["Email"].strip() != "" else vcards_noemail_folder
                        try:
                            vcard_path = download_vcard(data["VCardUrl"], folder_to_use, data["EntryId"])
                            master_entries[eid]["VCardPath"] = vcard_path
                        except Exception as ex:
                            print(f"vCard download failed for {data['VCardUrl']}: {ex}")
            else:
                print(f"Adding new entry {eid}.")
                if data["VCardUrl"]:
                    folder_to_use = vcards_email_folder if data["Email"].strip() != "" else vcards_noemail_folder
                    try:
                        vcard_path = download_vcard(data["VCardUrl"], folder_to_use, data["EntryId"])
                        data["VCardPath"] = vcard_path
                    except Exception as ex:
                        print(f"vCard download failed for {data['VCardUrl']}: {ex}")
                master_entries[eid] = data
        
        pos += len(entries)
        if pos > total_results:
            break
        
        human_sleep(2.0)
    
    # Convert master_entries dict to a list sorted by Company (for convenience)
    all_entries = sorted(master_entries.values(), key=lambda x: x.get("Company", "").lower())
    
    # Write/update the master CSV
    master_fieldnames = ["EntryId", "Updated", "Company", "Firstname", "Lastname", "Address", "Phone", "Email", "VCardPath"]
    write_csv(master_csv, master_fieldnames, all_entries)
    print(f"\nMaster CSV updated with {len(all_entries)} entries at: {master_csv}")
    
    # Filter entries into those with and without emails
    filtered_fieldnames = ["Company", "Firstname", "Lastname", "Address", "Phone", "Email", "VCardPath"]
    with_email = [entry for entry in all_entries if entry.get("Email", "").strip() != ""]
    without_email = [entry for entry in all_entries if entry.get("Email", "").strip() == ""]
    
    write_csv(csv_with_email, filtered_fieldnames, with_email)
    write_csv(csv_without_email, filtered_fieldnames, without_email)
    
    print(f"CSV with emails saved to: {csv_with_email}")
    print(f"CSV without emails saved to: {csv_without_email}")
    print("Done!")

if __name__ == "__main__":
    main()
