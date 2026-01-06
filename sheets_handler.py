import os
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

# Scope required for gspread
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_service_account_file():
    # Look for google sheets credentials in env or default to root
    return os.getenv("GOOGLE_SHEETS_JSON", "credentials.json")

def connect_to_sheet():
    """Connects to Google Sheets and returns the worksheet."""
    creds_file = get_service_account_file()
    
    if not os.path.exists(creds_file):
        raise FileNotFoundError(f"Google Service Account file not found at: {creds_file}")

    creds = Credentials.from_service_account_file(creds_file, scopes=SCOPES)
    client = gspread.authorize(creds)
    
    sheet_name = os.getenv("GOOGLE_SHEETS_NAME", "Tracking_Fullstackverse")
    # Open the spreadsheet
    try:
        sh = client.open(sheet_name)
    except gspread.SpreadsheetNotFound:
        raise ValueError(f"Spreadsheet '{sheet_name}' not found. Please share it with the service account email.")

    # Select the specific tab "BD Manager"
    try:
        worksheet = sh.worksheet("BD Manager")
    except gspread.WorksheetNotFound:
        # Fallback or create? For now let's error or default to first
        print("Tab 'BD Manager' not found, using first sheet.")
        worksheet = sh.sheet1
        
    return worksheet

def get_existing_data(worksheet):
    """
    Returns a set of existing unique identifiers (phones and names) to avoid duplicates.
    """
    try:
        # Get all records
        records = worksheet.get_all_records()
        existing_phones = set()
        existing_names = set()
        
        for row in records:
            # Assuming keys match header names if get_all_records uses them
            # Or we can just grab by index if we used get_all_values
            # Let's use get_all_values for safety as headers might vary slightly
            pass
            
        # Re-doing with get_all_values for robustness
        rows = worksheet.get_all_values()
        if len(rows) < 2:
            return set(), set()
            
        # Headers are row 0. Find indices.
        headers = [h.strip().lower() for h in rows[0]]
        
        idx_name = -1
        idx_phone = -1
        
        for i, h in enumerate(headers):
            if "name" in h: idx_name = i
            if "contact" in h or "phone" in h: idx_phone = i
            
        for row in rows[1:]:
            if idx_phone != -1 and len(row) > idx_phone:
                p = "".join(filter(str.isdigit, str(row[idx_phone])))
                if p: existing_phones.add(p)
            
            if idx_name != -1 and len(row) > idx_name:
                n = str(row[idx_name]).strip().lower()
                if n: existing_names.add(n)
                
        return existing_phones, existing_names
        
    except Exception as e:
        print(f"Warning: Could not fetch existing data for deduplication: {e}")
        return set(), set()

def append_lead(worksheet, data):
    """
    Appends a lead to the worksheet.
    Expected data dict keys: 'name', 'phone', 'profession', 'email', 'website', 'address', 'query'
    Columns mapping based on user request:
    A: Name of Lead
    B: Contact number of lead
    C: Nature/Profession of Lead
    D: Status (Default: 'New')
    E: Email
    F: Website (Extra)
    G: Address (Extra)
    H: Source Query (Extra)
    """
    row = [
        data.get("name", ""),
        data.get("phone", ""),
        data.get("profession", ""), # Mapped from query category or scraped info
        "New",                      # Status
        data.get("email", "")
    ]
    
    # Logic to find the first empty row (filling gaps) instead of appending to the very end
    # Get all values in Column A
    col_a_values = worksheet.col_values(1)
    
    try:
        # Check if there's a gap (an empty string) in the list of values
        # We start looking from index 1 (Row 2 equivalent) because index 0 is Header
        first_empty_idx = col_a_values.index("") 
        next_row = first_empty_idx + 1
    except ValueError:
        # No empty strings found in the middle, so append after the last value
        next_row = len(col_a_values) + 1
        
    # Update the specific row
    # range_name example: "A2"
    worksheet.update(range_name=f"A{next_row}", values=[row])
