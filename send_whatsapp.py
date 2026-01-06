import os
import time
import random
import pandas as pd
import pywhatkit as pwk
from dotenv import load_dotenv
import sheets_handler

# ---------- Helpers ----------
def clean_phone(s: str) -> str:
    s = str(s or "").strip()
    out = "".join(ch for ch in s if ch.isdigit() or ch == "+")
    if out.count("+") > 1:
        out = out.replace("+", "")
    return out

# ---------- Load config ----------
load_dotenv()
MSG_TEXT = os.getenv("MSG_TEXT", "Hello, I found your business online.")
MIN_DELAY = int(os.getenv("MIN_DELAY_SEC", "4"))
MAX_DELAY = int(os.getenv("MAX_DELAY_SEC", "9"))

def country_to_code(cc: str) -> str:
    mapping = {"IN": "+91", "US": "+1", "GB": "+44", "AE": "+971"}
    cc = (cc or "").upper().strip()
    return mapping.get(cc, "+")

DEFAULT_COUNTRY = os.getenv("DEFAULT_COUNTRY", "IN")
country_code_prefix = country_to_code(DEFAULT_COUNTRY)

def with_cc(n: str) -> str:
    if n.startswith("+"):
        return n
    return (country_code_prefix if country_code_prefix != "+" else "+91") + n

# ---------- Main ----------
def main():
    print("Connecting to Google Sheets...")
    try:
        worksheet = sheets_handler.connect_to_sheet()
    except Exception as e:
        print(f"Error connecting to sheets: {e}")
        return

    # Read all data
    print("Fetching records...")
    all_values = worksheet.get_all_values()
    if not all_values:
        print("Sheet is empty.")
        return
        
    headers = all_values[0]
    rows = all_values[1:]
    
    # Identify key columns by index helper
    def get_col_idx(name):
        try:
            return headers.index(name)
        except ValueError:
            return -1

    idx_phone = get_col_idx("Contact number of lead")
    idx_status = get_col_idx("Status")
    
    if idx_phone == -1 or idx_status == -1:
        print("Could not find 'Contact number of lead' or 'Status' columns.")
        return

    print(f"Found {len(rows)} rows. Filtering for 'New'...")
    
    count_sent = 0
    
    # Iterate rows (1-based index in Sheet logic, so start=2)
    for i, row in enumerate(rows, start=2):
        if i > len(rows) + 1: break # Safety
        
        # Safe access to columns
        phone_raw = row[idx_phone] if len(row) > idx_phone else ""
        status = row[idx_status] if len(row) > idx_status else ""
        
        if status.strip().lower() != "new":
            continue
            
        # Validate Phone
        phone_clean = clean_phone(phone_raw)
        if len(phone_clean) < 5:
            print(f"Row {i}: Invalid phone '{phone_raw}'. Marking Invalid.")
            worksheet.update_cell(i, idx_status + 1, "Invalid Phone")
            continue
            
        phone_final = with_cc(phone_clean)
        
        print(f"Row {i}: Sending to {phone_final}...")
        
        try:
            pwk.sendwhatmsg_instantly(
                phone_no=phone_final,
                message=MSG_TEXT,
                wait_time=15,
                tab_close=True,
                close_time=3
            )
            count_sent += 1
            # Update Status
            worksheet.update_cell(i, idx_status + 1, "Sent")
            print("   ✅ Sent & Updated")
            
        except Exception as e:
            print(f"   ❌ Failed: {e}")
            worksheet.update_cell(i, idx_status + 1, f"Error: {str(e)}")
            
        # Wait
        time.sleep(random.randint(MIN_DELAY, MAX_DELAY))
        
    print(f"Done. Messages sent: {count_sent}")

if __name__ == "__main__":
    main()
