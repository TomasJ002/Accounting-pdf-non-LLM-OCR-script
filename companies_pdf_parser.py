import os
import re
import pandas as pd
import pytesseract
import cv2
import numpy as np
import unicodedata
from thefuzz import fuzz
from pdf2image import convert_from_path

TESSERACT_PATH = r'C:\Program Files\TesseractOCR\tesseract.exe'
POPPLER_PATH = r'C:\poppler\Library\bin' 

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

SEARCH_CATEGORIES = {
    "Assets (Net)": ["aktiva celkem", "bilancni suma"],
    "Equity": ["vlastni kapital"],
    "Liabilities": ["cizi zdroje", "zavazky celkem"],
    "Revenue / Turnover": ["cisty obrat", "trzby za prodej zbozi", "trzby z prodeje vyrobku a sluzeb", "vykony"]
}

def strip_accents(text):
    if not text: return ""
    return unicodedata.normalize('NFKD', str(text)).encode('ASCII', 'ignore').decode('utf-8').lower()

def extract_numbers(line_text):
    if not line_text: return []
    matches = re.findall(r'-?\s*\d{1,3}(?:\s\d{3})*|-?\d+', str(line_text))
    numbers = []
    
    for n in matches:
        cleaned = n.replace(' ', '')
        try:
            numbers.append(int(cleaned))
        except:
            pass
            
    if numbers and len(str(abs(numbers[0]))) <= 3:
        numbers = numbers[1:]
        
    return numbers

def deep_analysis(text):
    data = {
        "Year": None,
        "Assets (Net)": None,
        "Equity": None,
        "Liabilities": None,
        "Revenue / Turnover": None
    }
    
    year_match = re.search(r'31\.12\.(\d{4})', text)
    if year_match:
        data["Year"] = year_match.group(1)

    FUZZY_THRESHOLD = 85 

    lines = text.split('\n')
    for i, line in enumerate(lines):
        line_clean = strip_accents(line)
        
        for category, keywords in SEARCH_CATEGORIES.items():
            if data[category] is None:
                
                match_found = False
                for kw in keywords:
                    similarity = fuzz.partial_ratio(kw, line_clean)
                    if similarity >= FUZZY_THRESHOLD:
                        match_found = True
                        break
                
                if match_found:
                    c = extract_numbers(line)
                    
                    if not c and i + 1 < len(lines):
                        c = extract_numbers(lines[i+1])
                        
                    if not c:
                        continue
                    
                    if category == "Assets (Net)":
                        raw_val = c[2] if len(c) >= 3 else c[0]
                    else:
                        raw_val = c[0]
                    
                    data[category] = raw_val * 1000

    return data

def process_pdf_full_ocr(file_path):
    file_name = os.path.basename(file_path)
    print(f"{file_name}", end="", flush=True)
    
    if not os.path.exists(TESSERACT_PATH):
        print("Tesseract not found!")
        return {"Status": "Error: Tesseract path", "File": file_name}
        
    try:
        images = convert_from_path(file_path, dpi=300, poppler_path=POPPLER_PATH)
        text = ""
        custom_config = r'--oem 3 --psm 6 -l ces+slk'
        
        for img in images:
            img_cv = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
            
            gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
            _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            kernel = np.ones((1, 1), np.uint8)
            clean_img = cv2.dilate(thresh, kernel, iterations=1)
            clean_img = cv2.erode(clean_img, kernel, iterations=1)
            
            text += pytesseract.image_to_string(clean_img, config=custom_config) + "\n"
            
    except Exception as e:
        print(f"OCR Error: {e}")
        return {"Status": f"OCR Error: {e}", "File": file_name}

    results = deep_analysis(text)
    results["Status"] = "Full OCR"
    results["File"] = file_name
    
    results["ICO"] = file_name.split('_')[0] 
    
    print(f" -> Done")
    return results

if __name__ == "__main__":
    FOLDER = "zavierky_pdf"
    OUTPUT_EXCEL = "firmy_ares_parsed.xlsx"
    
    if not os.path.exists(FOLDER):
        print(f"Folder '{FOLDER}' does not exist!")
    else:
        all_data = []
        pdf_list = [f for f in os.listdir(FOLDER) if f.lower().endswith('.pdf')]
        print(f"Amount of documents: {len(pdf_list)}")
        
        for i, f in enumerate(pdf_list):
            data = process_pdf_full_ocr(os.path.join(FOLDER, f))
            all_data.append(data)
            
            if (i + 1) % 5 == 0:
                df_temp = pd.DataFrame(all_data)
                try:
                    df_temp.to_excel(OUTPUT_EXCEL, index=False)
                except PermissionError:
                    df_temp.to_excel(OUTPUT_EXCEL.replace('.xlsx', '_backup.xlsx'), index=False)

        df = pd.DataFrame(all_data)
        column_order = ["ICO", "Year", "Status", "Revenue / Turnover", "Assets (Net)", "Equity", "Liabilities", "File"]
        existing_columns = [c for c in column_order if c in df.columns]
        
        try:
            df[existing_columns].to_excel(OUTPUT_EXCEL, index=False)
            print(f"Results in: {OUTPUT_EXCEL}")
        except PermissionError:
            backup_name = OUTPUT_EXCEL.replace('.xlsx', '_FINAL_BACKUP.xlsx')
            df[existing_columns].to_excel(backup_name, index=False)
            print(f"{OUTPUT_EXCEL} was open in Excel. Saved to: {backup_name}")