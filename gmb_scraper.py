import os, re, time, random, json, urllib.parse
import config
import sheets_handler
from playwright.sync_api import sync_playwright

# ── Configuration ─────────────────────────────────────────────────────────────
HEADLESS = os.getenv("HEADLESS", "false").lower() in ("1", "true", "yes")
MAX_RESULTS_PER_KEYWORD = int(os.getenv("MAX_RESULTS", "10")) # Default lower for deep scraping

# ── Helpers ───────────────────────────────────────────────────────────────────
def human_delay(a=1.0, b=2.0):
    time.sleep(random.uniform(a, b))

def clean_text(text):
    return (text or "").strip()

def extract_emails(text):
    return list(set(re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text or "")))

def extract_phones(text):
    # Basic phone extraction, can be improved
    return list(set(re.findall(r"\+?\d[\d -]{8,12}\d", text or "")))

def is_relevant(text):
    text_lower = text.lower()
    return any(k.lower() in text_lower for k in config.RELEVANCE_KEYWORDS)

def find_decision_makers(text):
    found = []
    lines = text.split('\n')
    for line in lines:
        for title in config.DECISION_MAKERS:
            if title.lower() in line.lower():
                # Try to extract a name near the title? 
                # For now, just capturing the context line
                if len(line.strip()) < 100: # detailed enough
                    found.append(line.strip())
    return list(set(found))

# ── Deep Scraping Logic ───────────────────────────────────────────────────────
def deep_scrape_website(context, url):
    print(f"   → Deep scraping: {url}")
    meta = {
        "emails": [],
        "phones": [],
        "decision_makers": [],
        "relevant": False
    }
    
    try:
        page = context.new_page()
        page.goto(url, timeout=15000)
        human_delay(2, 3)
        
        # 1. Analyze Home Page
        content = page.content()
        text = page.inner_text("body")
        
        if is_relevant(text):
            meta["relevant"] = True
            
        meta["emails"].extend(extract_emails(content))
        meta["phones"].extend(extract_phones(text))
        meta["decision_makers"].extend(find_decision_makers(text))
        
        # 2. Find Sub-pages (About, Team, Contact, Careers)
        # Look for links
        links = page.query_selector_all("a")
        sub_pages = set()
        keywords = ["about", "team", "contact", "career", "leadership", "faculty", "staff"]
        
        for link in links:
            href = link.get_attribute("href")
            if href and any(k in href.lower() for k in keywords):
                # Resolve relative URLs
                full_url = urllib.parse.urljoin(url, href)
                if url in full_url: # Stay on same domain
                    sub_pages.add(full_url)
        
        # Limit sub-pages to visit
        for sub_url in list(sub_pages)[:3]: # Visit max 3 sub-pages to save time
            try:
                print(f"     → Visiting sub-page: {sub_url}")
                page.goto(sub_url, timeout=10000)
                human_delay(1, 2)
                sub_content = page.content()
                sub_text = page.inner_text("body")
                
                meta["emails"].extend(extract_emails(sub_content))
                meta["phones"].extend(extract_phones(sub_text))
                meta["decision_makers"].extend(find_decision_makers(sub_text))
            except Exception:
                continue

        page.close()
        
    except Exception as e:
        print(f"   ⚠ Error deep scraping {url}: {e}")
        try:
             page.close()
        except:
            pass
            
    # Dedupe
    meta["emails"] = list(set(meta["emails"]))
    meta["phones"] = list(set(meta["phones"]))
    meta["decision_makers"] = list(set(meta["decision_makers"]))
    
    return meta

# ── Main Maps Scraper ─────────────────────────────────────────────────────────
def run_scraper():
    # Connect to Sheets first
    try:
        worksheet = sheets_handler.connect_to_sheet()
        print("✅ Connected to Google Sheets")
        
        # Load existing for dedupe
        print("⏳ Fetching existing leads for deduplication...")
        existing_phones, existing_names = sheets_handler.get_existing_data(worksheet)
        print(f"   found {len(existing_phones)} phones and {len(existing_names)} names to skip.")
        
    except Exception as e:
        print(f"❌ Failed to connect to Sheets: {e}")
        return

    # 1. Pre-generate ALL combinations (Do this BEFORE launching browser to save CPU)
    print("🎲 Generating search combinations...")
    start_gen = time.time()
    all_combinations = []
    
    # Get Locations cleanly
    raw_locations = getattr(config, "LOCATIONS", [])
    # Filter out empty strings and None, but ensure at least one empty string if list is effectively empty
    locations = [str(l).strip() for l in raw_locations if l]
    if not locations:
        locations = [""]
    
    count_c = 0
    for category, keywords in config.SEARCH_KEYWORDS.items():
        count_c += 1
        # print(f"   Debug: Processing Category {count_c} ({category})...") 
        for keyword in keywords:
                for loc in locations:
                    query = f"{keyword} in {loc}" if loc else keyword
                    all_combinations.append((category, query))
    
    end_gen = time.time()
    print(f"   ⏱️  Generation took {end_gen - start_gen:.4f} seconds.")
    
    # Shuffle
    random.shuffle(all_combinations)
    print(f"🎲 Randomized {len(all_combinations)} search queries to process.")

    with sync_playwright() as p:
        print("🚀 Launching Browser (This may take a moment)...")
        browser = p.chromium.launch(headless=HEADLESS)
        context = browser.new_context()
        
        # 3. Process the randomized list
        for category, query in all_combinations:
            print(f"\n═══ Processing: {query} [{category}] ═══")
            
            try:
                    page = context.new_page()
                    search_url = "https://www.google.com/maps/search/?q=" + urllib.parse.quote_plus(query) + "&hl=en"
                    page.goto(search_url)
                    human_delay(2, 3)
                    
                    # Consent check
                    try:
                        # Use query_selector check instead of click for speed if not present
                        if page.query_selector('button[aria-label="Accept all"]'):
                            page.click('button[aria-label="Accept all"]', timeout=2000)
                    except:
                        pass


                    # Scroll
                    for _ in range(5): # Scroll a few times
                        page.evaluate("window.scrollBy(0, 2000)")
                        human_delay(0.5, 1)

                    # Get Cards
                    anchors = page.query_selector_all('a[href*="google.com/maps/place"]')
                    hrefs = list(set([a.get_attribute("href") for a in anchors if a.get_attribute("href")]))
                    print(f"Found {len(hrefs)} initial results")
                    
                    processed_count = 0
                    for maps_url in hrefs[:MAX_RESULTS_PER_KEYWORD]:
                        try:
                            # Optimization: We can't easily check name/phone without visiting at least the GMB card.
                            # But we can check if the Maps URL or something was saved. 
                            # Since we only saved Name/Phone, we must visit the card to extract them first.
                            
                            page.goto(maps_url)
                            human_delay(1, 2)
                            
                            # Basic Info
                            name = clean_text(page.inner_text('h1') if page.query_selector('h1') else "")
                            
                            # DEDUPE BY NAME EARLY
                            if name.lower() in existing_names:
                                print(f"   (Skip) Duplicate Name: {name}")
                                continue

                            # Phone
                            phone_el = page.query_selector('button[data-item-id^="phone:tel:"]')
                            phone = clean_text(phone_el.get_attribute("aria-label").replace("Phone:", "") if phone_el else "")
                            
                            # DEDUPE BY PHONE
                            clean_p = "".join(filter(str.isdigit, phone))
                            if clean_p and clean_p in existing_phones:
                                print(f"   (Skip) Duplicate Phone: {phone}")
                                continue

                            # Website
                            website_el = page.query_selector('a[data-item-id="authority"]')
                            website = website_el.get_attribute("href") if website_el else ""
                            
                            print(f"Found: {name} | Phone: {phone} | Web: {website}")
                            
                            # DEEP SCRAPE
                            extra_data = {"emails": [], "relevant": True} # Default relevant if no website to check
                            if website:
                                extra_data = deep_scrape_website(context, website)

                            if not extra_data["relevant"]:
                                print("   → Skipping: Low relevance")
                                continue
                                
                            # Prepare Data for Sheet
                            # Priority: Scraped Email > None
                            # Priority: Scraped Phone > GMB Phone
                            
                            final_email = extra_data["emails"][0] if extra_data["emails"] else ""
                            final_phone = phone
                            if not final_phone and extra_data["phones"]:
                                final_phone = extra_data["phones"][0]
                                
                            lead_data = {
                                "name": name,
                                "phone": final_phone,
                                "profession": category, # e.g., STUDENT_DEGREE
                                "email": final_email,
                                "website": website,
                                "query": query,
                                "address": ""
                            }
                            
                            # Append to Sheet
                            sheets_handler.append_lead(worksheet, lead_data)
                            print("   ✅ Saved to Sheet")
                            processed_count += 1
                            
                        except Exception as e:
                            print(f"Error processing card: {e}")
                            continue

                    page.close()
                    
            except Exception as e:
                print(f"Error processing query '{query}': {e}")
                try: page.close()
                except: pass
        
        browser.close()

if __name__ == "__main__":
    run_scraper()
