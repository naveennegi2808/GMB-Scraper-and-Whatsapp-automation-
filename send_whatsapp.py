import os
import time
import random
import urllib.parse
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import sheets_handler

# ---------- Load config ----------
load_dotenv()
# MSG_TEXT = os.getenv("MSG_TEXT", "Hello, I found your business online.") # <-- Replaced by variations below

MESSAGE_VARIATIONS = [
    """Hi! Reaching out from SkillVerse.
We run a high-value 'AI Agent Workshop' where students build real AI agents (No coding required).
Link: https://fullstackverse.com/skillverse/agent-ai-workshop

We are looking for partners for collaboration. You can earn a commission for every enrollment.
Let me know if you are interested in a profitable partnership! 🤝""",

    """Hello 👋 this is Business Development Manager from SkillVerse.
We are hosting a live AI workshop where participants get a Premium Certificate and Lifetime Access.
Check it out: https://fullstackverse.com/skillverse/agent-ai-workshop

We offer a commission-based partnership where you earn for every student that joins.
Let me know if you are interested in collaborating.""",

    """Hi, hope you are doing well.
We conduct premium AI workshops covering Zero-Coding AI Agents & Automation.
Details: https://fullstackverse.com/skillverse/agent-ai-workshop

I see a great opportunity for a partnership. You will receive a commission for every participant referred by you.
If a collaboration sounds interesting, please reply!""",

    """Hi! We run professional AI webinars at SkillVerse, teaching Step-by-Step AI Agent building.
Workshop: https://fullstackverse.com/skillverse/agent-ai-workshop

We can have a partnership where you earn a commission for promoting our workshop.
If you are interested in this revenue-sharing opportunity, let me know.""",

    """Hello! Reaching out regarding a potential collaboration.
We provide high-quality workshops on Building AI Agents (100% Practical).
Link: https://fullstackverse.com/skillverse/agent-ai-workshop

We are looking for partners and you will get a commission on every booking.
Let me know if you'd like to partner with us!"""
]

MIN_DELAY = int(os.getenv("MIN_DELAY_SEC", "4"))
MAX_DELAY = int(os.getenv("MAX_DELAY_SEC", "9"))
HEADLESS = os.getenv("HEADLESS", "false").lower() in ("1", "true", "yes")

# ---------- Helpers ----------
def clean_phone(s: str) -> str:
    s = str(s or "").strip()
    out = "".join(ch for ch in s if ch.isdigit() or ch == "+")
    if out.count("+") > 1:
        out = out.replace("+", "")
    return out

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

def human_delay(a=MIN_DELAY, b=MAX_DELAY):
    time.sleep(random.randint(a, b))

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
    
    # Identify key columns
    def get_col_idx(name):
        try: return headers.index(name)
        except ValueError: return -1

    idx_phone = get_col_idx("Contact number of lead")
    idx_status = get_col_idx("Status")
    
    if idx_phone == -1 or idx_status == -1:
        print("Could not find 'Contact number of lead' or 'Status' columns.")
        return

    print(f"Found {len(rows)} rows. Filtering for 'New'...")
    
    # User Input for Limit
    target_count = len([r for r in rows if len(r) > idx_status and r[idx_status].strip().lower() == "new"])
    print(f"   (There are {target_count} 'New' leads available to message)")
    
    try:
        user_limit = input(f"🎯 How many people do you want to message? (Press Enter for all {target_count}): ").strip()
        if not user_limit:
            max_messages = target_count
        else:
            max_messages = int(user_limit)
    except ValueError:
        print("Invalid input. Defaulting to ALL.")
        max_messages = target_count
    
    print(f"   → Will stop after sending {max_messages} messages.\n")
    
    # ─── PLAYWRIGHT SETUP ─────────────────────────────────────────────────────────
    # We use a persistent context so the user stays logged in (QR code saved).
    user_data_dir = os.path.join(os.getcwd(), "whatsapp_session")
    
    print(f"🚀 Launching WhatsApp Automation (Session: {user_data_dir})")
    print("ℹ️  Note: If this is your first time, you will need to scan the QR code in the browser window.")
    
    with sync_playwright() as p:
        # Launch persistent context (Headless=False always recommended for WhatsApp to scan QR)
        browser = p.chromium.launch_persistent_context(
            user_data_dir,
            headless=False, # Must be visible to authenticate
            channel="chrome", # Use installed Chrome if available, else chromium
            args=["--start-maximized", "--no-sandbox"]
        )
        
        page = browser.pages[0]
        if not page: page = browser.new_page()
        
        # Open WhatsApp Web once
        page.goto("https://web.whatsapp.com/")
        
        print("⏳ Waiting for login... (Please scan QR code if needed)")
        try:
            # Wait for the chat list side panel (indicates successful login)
            # Increased timeout to 300 seconds (5 minutes) based on user request
            page.wait_for_selector('div[id="pane-side"]', timeout=300000) 
            print("✅ Logged in successfully!")
        except:
            print("❌ Timeout waiting for login. Please try running again and scanning quickly.")
            browser.close()
            return
            
        count_sent = 0
        
        # Track sent numbers in this session to avoid duplicates
        session_sent_files = set()
        last_chat_header = None
        
        # Iterate rows (start=2 because row 1 is header)
        for i, row in enumerate(rows, start=2):
            if i > len(rows) + 1: break
            
            # CHECK LIMIT
            if count_sent >= max_messages:
                print(f"\n🛑 Reached limit of {max_messages} messages. Stopping.")
                break
            
            phone_raw = row[idx_phone] if len(row) > idx_phone else ""
            status = row[idx_status] if len(row) > idx_status else ""
            
            if status.strip().lower() != "new":
                continue
                
            # Validate Phone
            # Remove leading zeros which break deep linking when combined with country code
            # e.g. 098... becomes 98...
            phone_clean = clean_phone(phone_raw).lstrip('0')
            
            if len(phone_clean) < 5:
                print(f"Row {i}: Invalid phone '{phone_raw}'. Marking Invalid.")
                worksheet.update_cell(i, idx_status + 1, "Invalid Phone")
                continue
            
            # --- DEDUPLICATION CHECK ---
            if phone_clean in session_sent_files:
                print(f"Row {i}: Duplicate number {phone_clean}. Skipping.")
                worksheet.update_cell(i, idx_status + 1, "Duplicate-Skip")
                continue
                
            phone_final = with_cc(phone_clean)
            
            print(f"Row {i}: Sending to {phone_final}...")
            
            try:
                # 0. RESET STATE: Close current chat to ensure a fresh navigation
                try: 
                    page.keyboard.press("Escape")
                    time.sleep(1.0)
                except: pass

                # 1. Select a Random Message
                chosen_msg = random.choice(MESSAGE_VARIATIONS)
                encoded_msg = urllib.parse.quote(chosen_msg)
                
                url = f"https://web.whatsapp.com/send?phone={phone_final}&text={encoded_msg}"
                
                # RETRY LOOP for Navigation
                # If first try fails or hangs, we try once more
                for attempt in range(2):
                    try:
                        print(f"   🚀 Navigating to {phone_final} (Attempt {attempt+1})...")
                        # 'domcontentloaded' is faster for SPAs like WhatsApp
                        page.goto(url, wait_until='domcontentloaded', timeout=15000)
                        
                        # Wait for context switch
                        time.sleep(5.0) 
                        
                        # Check if it looks like it worked (Input box exists?)
                        input_box_selector = 'div[contenteditable="true"][role="textbox"]'
                        input_generic = 'div[contenteditable="true"]'
                        
                        if page.query_selector(input_box_selector) or page.query_selector(input_generic):
                             print("   ✅ Chat loaded.")
                             break # Success!
                        else:
                             print("   ⚠️  Chat didn't load yet...")
                             if attempt == 0:
                                 print("   🔄 Retrying navigation...")
                                 continue
                    except Exception as e:
                        print(f"   ⚠️  Nav error: {e}")
                
                # 2. Wait for the chat to load (Input Box or Invalid Popup)
                try:
                    # Wait for the input box
                    # FIX: Use specific selector inside #main to avoid conflict with Search box
                    input_box_selector = '#main footer div[contenteditable="true"][role="textbox"]'
                    try:
                        page.wait_for_selector(input_box_selector, timeout=20000)
                    except Exception:
                        try:
                            # Fallback 1: Generic footer editable
                            input_box_selector = '#main footer div[contenteditable="true"]' 
                            page.wait_for_selector(input_box_selector, timeout=5000)
                        except:
                            # Fallback 2: Aria label based (very specific)
                            try:
                                input_box_selector = 'div[aria-placeholder="Type a message"]'
                                page.wait_for_selector(input_box_selector, timeout=5000)
                            except:
                                if page.query_selector('div[data-animate-modal-popup="true"]'):
                                    print("   ⚠️  Invalid WhatsApp Number (Popup Detected).")
                                    worksheet.update_cell(i, idx_status + 1, "Invalid WhatsApp")
                                    continue
                                else:
                                    raise Exception("Timeout waiting for chat to load (input box not found)")
                    
                    # --- ROBUST SAFETY CHECK: Stale Header Detection ---
                    # If the header hasn't changed since the last message, WE ARE STUCK.
                    header_text = ""
                    try:
                        header_loc = page.locator("#main header")
                        header_loc.wait_for(timeout=5000)
                        header_text = header_loc.inner_text().strip()
                        print(f"   ℹ️  Header: '{header_text.replace(chr(10), ' ')}'")
                        
                        # 1. EMPTY Check
                        if not header_text:
                             print("   🛑 STOP: Header is empty. Navigation failed.")
                             worksheet.update_cell(i, idx_status + 1, "Nav-Error-Empty")
                             continue
                             
                        # 2. STALE Check (Did we change chats?)
                        if last_chat_header and header_text == last_chat_header:
                             print(f"   🛑 STOP: Header identical to previous contact ('{header_text}'). Navigation failed.")
                             worksheet.update_cell(i, idx_status + 1, "Nav-Error-Stuck")
                             continue
                             
                        # 3. SPECIFIC BLACKLIST Check (Double safety)
                        if "Achievers" in header_text or "4Achievers" in header_text:
                             print(f"   🛑 STOP: Blacklisted header detected.")
                             worksheet.update_cell(i, idx_status + 1, "Nav-Error-Stuck")
                             continue
                        
                        # 4. PHONE NUMBER MISMATCH Check
                        # If the header contains a '+' followed by digits, it might be a phone number.
                        # We check if it matches our target.
                        # Extract all phone-like patterns from header
                        import re
                        header_phones = re.findall(r'\+\d[\d\s]+', header_text)
                        if header_phones:
                            # Clean them up for comparison
                            header_phones_clean = [clean_phone(p) for p in header_phones]
                            target_phone_clean = clean_phone(phone_final)
                            
                            # If we found phones, and NONE of them match our target
                            if header_phones_clean and target_phone_clean not in header_phones_clean:
                                # Safe check: sometimes header has the USER'S phone (unlikely) or other info.
                                # But usually, if it shows a number, it's the contact's number.
                                print(f"   🛑 STOP: Header phone {header_phones_clean} != Target {target_phone_clean}. Wrong Chat.")
                                worksheet.update_cell(i, idx_status + 1, "Nav-Error-WrongChat")
                                continue

                        # Update last header
                        last_chat_header = header_text
                        
                    except Exception as e:
                        print(f"   ⚠️  Header check failed: {e}. Aborting for safety.")
                        worksheet.update_cell(i, idx_status + 1, "Nav-Error-CheckFail")
                        continue

                    # Small human pause before sending
                    time.sleep(random.uniform(1.0, 3.0))
                    
                    # 3. WAKE UP THE UI: Focus and ensure Send button becomes ready
                    try:
                        inp = page.locator(input_box_selector)
                        inp.click()
                        time.sleep(0.5)
                        # Type space backspace to trigger events
                        inp.type(" ", delay=100)
                        inp.press("Backspace")
                        time.sleep(1.0) # Wait for UI to react
                    except: pass

                    # 4. ROBUST SEND: Wait for Send Button and Press Enter
                    try:
                        # Attempt to find the Send button
                        send_selectors = [
                            'span[data-icon="send"]',
                            'button[aria-label="Send"]',
                            'div[aria-label="Send"]',
                            'span[data-testid="send"]'
                        ]
                        
                        send_btn_loc = None
                        for sel in send_selectors:
                            if page.locator(sel).is_visible():
                                send_btn_loc = page.locator(sel)
                                break
                        
                        # Primary Send Method: PRESS ENTER on Input
                        # We do this even if we found the button, it's more human-like.
                        # But we check if the button disappears (meaning sent)
                        page.locator(input_box_selector).press("Enter")
                        time.sleep(1.0)
                        
                        # Fallback: Check if Send Button still exists. If so, CLICK IT.
                        if send_btn_loc and send_btn_loc.is_visible():
                            print(f"   ℹ️  'Enter' didn't work. Clicking Send button...")
                            send_btn_loc.click()
                            time.sleep(1.0)
                            
                    except Exception as e:
                        print(f"   ⚠️  Send action error: {e}")
                    
                    # 5. FINAL VERIFICATION: Is Input Box Empty?
                    try:
                        input_el = page.query_selector(input_box_selector)
                        if input_el:
                            content = input_el.text_content()
                            if content and len(content.strip()) > 0:
                                print(f"   ❌ FAILED: Message text still remains in input box.")
                                worksheet.update_cell(i, idx_status + 1, "Draft-Error")
                                
                                # CRITICAL: Clear the box so it doesn't block next navigation
                                try:
                                    print("   🧹 Clearing stuck draft...")
                                    page.locator(input_box_selector).fill("")
                                except: pass
                                
                                continue
                    except: pass
                    
                    print("   ✅ Message Sent!")
                    
                    # Update Status and Session Cache
                    worksheet.update_cell(i, idx_status + 1, "Sent")
                    session_sent_files.add(phone_clean)
                    count_sent += 1
                    
                    # Human Delay after sending
                    wait_s = random.randint(MIN_DELAY, MAX_DELAY)
                    print(f"   💤 Waiting {wait_s}s...")
                    time.sleep(wait_s)
                    
                except Exception as e:
                    print(f"   ❌ Error sending: {e}")
                    worksheet.update_cell(i, idx_status + 1, "Error")

            except Exception as e:
                print(f"   ❌ Failed: {e}")
                worksheet.update_cell(i, idx_status + 1, f"Error: {str(e)}")
                
        print(f"🎉 Done. Messages sent: {count_sent}")
        time.sleep(3)
        browser.close()

if __name__ == "__main__":
    main()
