import requests
import time
import random
import csv
import os
import json
from datetime import datetime

API_KEY = "API--KEY--GOES--HERE"  # Your API key
BASE_URL = "https://tel.search.ch/api/"
USAGE_FILE = "api_usage.json"
MONTHLY_QUOTA = 1000

def human_sleep(base_delay):
    """
    Sleep for base_delay seconds with ±20% jitter.
    """
    actual_delay = base_delay * random.uniform(0.8, 1.2)
    time.sleep(actual_delay)

def load_api_usage():
    """
    Loads API usage info from a JSON file.
    If the saved month is not the current month, reset the counter.
    """
    current_month = datetime.now().strftime("%Y-%m")
    try:
        with open(USAGE_FILE, "r") as f:
            usage = json.load(f)
        # Reset counter if month has changed
        if usage.get("month") != current_month:
            usage = {"month": current_month, "calls": 0}
    except Exception:
        usage = {"month": current_month, "calls": 0}
    return usage

def save_api_usage(usage):
    """
    Saves API usage info to a JSON file.
    """
    with open(USAGE_FILE, "w") as f:
        json.dump(usage, f)

def increment_api_counter():
    """
    Increments the API call counter by one and saves it.
    Returns the new total.
    """
    usage = load_api_usage()
    usage["calls"] += 1
    save_api_usage(usage)
    return usage["calls"]

def show_remaining_api_calls():
    """
    Loads the usage info and prints how many API calls are remaining.
    """
    usage = load_api_usage()
    remaining = MONTHLY_QUOTA - usage.get("calls", 0)
    print(f"You have {remaining} API calls remaining for this month.")
    if remaining <= 0:
        print("WARNING: Your API quota has been exhausted!")
    return remaining

def get_user_input():
    what = input("Bitte gib den Suchbegriff (z.B. metallbau) ein: ").strip()
    where = input("Bitte gib den Ort/Kanton (z.B. AG) ein: ").strip()
    return what, where

def fetch_results(what, where, start=0, maxnum=10):
    """
    Fetch a batch of listings from the API with the given parameters.
    Increments the API call counter on a successful call.
    """
    usage = load_api_usage()
    if usage.get("calls", 0) >= MONTHLY_QUOTA:
        raise Exception("API quota exceeded for this month!")
        
    params = {
        "key": API_KEY,
        "what": what,
        "where": where,
        "start": start,
        "maxnum": maxnum,
        "lang": "de",    # optional (de, en, fr, it)
        "format": "json" # ensure JSON response
    }
    response = requests.get(BASE_URL, params=params)
    response.raise_for_status()
    # Count this API call
    increment_api_counter()
    return response.json()

def download_vcard(vcard_url, output_folder, entry_id):
    """
    Download vCard from vcard_url and save to output_folder with a unique filename.
    """
    if not vcard_url:
        return None
    
    print(f"Downloading vCard from: {vcard_url}")
    resp = requests.get(vcard_url)
    resp.raise_for_status()
    
    filename = os.path.basename(vcard_url)
    if not filename.lower().endswith(".vcf"):
        filename = f"{entry_id}.vcf"
    
    filepath = os.path.join(output_folder, filename)
    with open(filepath, 'wb') as f:
        f.write(resp.content)
    
    return filepath

def main():
    """
    Main routine:
      1. Show remaining API calls.
      2. Get search parameters.
      3. Query the API (with pagination) until all results are collected.
      4. Optionally download each listing's vCard.
      5. Save results in a CSV file.
    """
    remaining = show_remaining_api_calls()
    if remaining <= 0:
        return

    what, where = get_user_input()
    output_folder = f"results_{what}_{where}"
    os.makedirs(output_folder, exist_ok=True)
    
    csv_path = os.path.join(output_folder, "results.csv")
    fieldnames = ["Company", "Firstname", "Lastname", "Address", "Phone", "Email", "VCardPath"]
    
    all_entries = []
    start = 0
    maxnum = 10  # number of results per API call
    total_found = 0
    downloaded_count = 0

    while True:
        # Fetch a batch of up to 'maxnum' results
        try:
            data = fetch_results(what, where, start=start, maxnum=maxnum)
        except Exception as ex:
            print(f"Error during API call: {ex}")
            break
        
        # 'count' is the number of items in this batch; 'total' is overall total
        count = data.get("count", 0)
        total_found = data.get("total", 0)
        
        entries = data.get("entries", [])
        if not entries:
            break
        
        print(f"Fetched {count} entries, total fetched so far: {start + count}/{total_found}")

        for entry in entries:
            company = entry.get("company", "")
            firstname = entry.get("firstname", "")
            lastname = entry.get("lastname", "")
            
            # Combine address fields
            street = entry.get("street", "")
            streetno = entry.get("streetno", "")
            zip_ = entry.get("zip", "")
            city = entry.get("city", "")
            address = f"{street} {streetno}, {zip_} {city}".strip(", ")
            
            # Concatenate phone numbers if provided
            phone = ""
            if "phone" in entry:
                phone_list = [p.get("dial", "") for p in entry["phone"] if p.get("dial")]
                phone = ", ".join(phone_list)
            
            # Concatenate email addresses if provided
            email = ""
            if "email" in entry:
                email_list = [e for e in entry["email"] if e]
                email = ", ".join(email_list)
            
            # Download vCard if available
            vcard_url = entry.get("vcard", "")
            vcard_path = ""
            if vcard_url:
                try:
                    vcard_path = download_vcard(vcard_url, output_folder, entry_id=entry.get("id", "unknown"))
                    downloaded_count += 1
                except Exception as ex:
                    print(f"vCard download failed for {vcard_url}: {ex}")
            
            all_entries.append({
                "Company": company,
                "Firstname": firstname,
                "Lastname": lastname,
                "Address": address,
                "Phone": phone,
                "Email": email,
                "VCardPath": vcard_path
            })
        
        start += count
        if start >= total_found:
            break

        human_sleep(2.0)

    # Save all entries to CSV
    print(f"\nSaving {len(all_entries)} entries to {csv_path}")
    with open(csv_path, 'w', encoding='utf-8-sig', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in all_entries:
            writer.writerow(row)

    print(f"Done! Downloaded {downloaded_count} vCards (if available).")
    print(f"CSV saved to: {csv_path}")

if __name__ == "__main__":
    main()
