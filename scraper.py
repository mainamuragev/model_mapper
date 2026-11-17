import pandas as pd
from playwright.sync_api import sync_playwright
import time
import re
import os

INPUT_FILE = "input_parts.xlsx"
OUTPUT_FILE = "Develon_by_price_FRESH.xlsx"

def clean_models(text_list):
    """Extracts only the model names from the messy HTML string."""
    models = set()
    for text in text_list:
        # Regex to find standard Develon models (e.g., DX225LC-7, SD200, DL300A)
        raw_matches = re.findall(r'\b[A-Z]{2,3}\d{2,4}[A-Z0-9-]*\b', text)
        for match in raw_matches:
            # Filter out junk from the file paths and codes
            if not match.startswith(('FIG', 'PDF', 'EXP', 'HDI', 'CEB', 'KOR')):
                models.add(match)
    return ", ".join(sorted(models))

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"Error: Could not find '{INPUT_FILE}'.")
        return

    print("Loading Excel file...")
    xls = pd.ExcelFile(INPUT_FILE)
    
    with sync_playwright() as p:
        print("Launching Microsoft Edge...")
        # Using msedge instead of chromium to bypass the slow download
        browser = p.chromium.launch(headless=False, slow_mo=50, channel="msedge")
        page = browser.new_page()
        page.goto("https://gps.hd-ce.com/cepcn/a/info/parts_info_vw.jsp")

        print("\n" + "="*60)
        print("ACTION REQUIRED: Log in to the portal in the opened Edge window.")
        print("Once you are on the Parts Information screen, return here.")
        input("Press ENTER to start scraping...")
        print("="*60 + "\n")

        # Loop through every sheet in the Excel file
        for sheet_name in xls.sheet_names:
            print(f"\n--- Processing Sheet: {sheet_name} ---")
            
            df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
            
            if len(df.columns) < 5:
                print("Skipping sheet: Not enough columns.")
                continue

            for index, row in df.iterrows():
                # Column B (index 1) is "Part No (without DS)"
                part_no = str(row.iloc[1]).strip()

                if part_no == 'nan' or part_no == '':
                    continue

                print(f"[Row {index+2}] Searching Part No: {part_no}")
                
                try:
                    # 1. Search
                    search_box = page.locator("input[type='text']").first
                    search_box.fill("") 
                    search_box.fill(part_no)
                    page.locator("input[value='Search'], button:has-text('Search')").first.click()
                    
                    # 2. Wait for load
                    page.wait_for_load_state("networkidle")
                    time.sleep(1.5) 

                    # 3. Extract Models
                    applied_models = []
                    rows = page.locator("tr[onclick*='focusTextAndImageSearch']").all()
                    
                    if len(rows) > 0:
                        for r in rows:
                            onclick_text = r.get_attribute("onclick")
                            if onclick_text:
                                applied_models.append(onclick_text)
                        
                        final_model_string = clean_models(applied_models)
                        print(f"   -> Found: {final_model_string}")
                    else:
                        final_model_string = ""
                        print("   -> No models found.")

                    # 4. Overwrite Column E (index 4) with the new data
                    df.at[index, df.columns[4]] = final_model_string

                except Exception as e:
                    print(f"   -> ERROR searching {part_no}: {e}")
                
                time.sleep(2)

            # Save the updated sheet to the new Excel file
            print(f"Saving sheet '{sheet_name}' to {OUTPUT_FILE}...")
            if os.path.exists(OUTPUT_FILE) and sheet_name == xls.sheet_names[0]:
                os.remove(OUTPUT_FILE) 
                
            with pd.ExcelWriter(OUTPUT_FILE, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"\nSUCCESS! All sheets processed. Data saved to {OUTPUT_FILE}")
        browser.close()

if __name__ == "__main__":
    main()
