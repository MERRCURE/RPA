import os
import re
import time
import csv
import glob
import zipfile
import shutil
import platform  # Importiert, um das Betriebssystem zu prüfen
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, StaleElementReferenceException
)

try:
    from pdf2image import convert_from_path
    import pytesseract
except Exception:
    convert_from_path = None
    pytesseract = None

def get_poppler_path():
    """Findet automatisch den Poppler-Installationspfad von winget."""
    
    # --- START FALLBACK-LISTE (WICHTIGSTE ÄNDERUNG) ---
    # Feste Pfade für Poppler. Ersetzen Sie den ersten Eintrag durch Ihren Pfad!
    FALLBACK_PATHS = [
        # BITTE HIER DEN DURCH DIE MANUELLE SUCHE GEFUNDENEN PFAD EINFÜGEN!
        r"C:\Users\spenl\AppData\Local\Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe\poppler-25.07.0\Library\bin"
    ]

    for path in FALLBACK_PATHS:
        if os.path.isdir(path):
            print(f"INFO: Poppler-Pfad über Hardcode-Fallback erkannt: {path}")
            return path
    # --- ENDE FALLBACK-LISTE ---
    
    if platform.system() == "Windows":
        # Standard-Installationspfad von winget
        poppler_dirs = glob.glob(r"C:\Program Files\poppler-*")
        if poppler_dirs:
            # Nimm die neuste Version, falls mehrere vorhanden sind
            latest_poppler_dir = max(poppler_dirs, key=os.path.getctime)
            poppler_bin_path = os.path.join(latest_poppler_dir, "bin")
            if os.path.isdir(poppler_bin_path):
                print(f"INFO: Poppler-Pfad automatisch erkannt: {poppler_bin_path}")
                return poppler_bin_path
        
        print("WARNUNG: Poppler-Pfad nicht automatisch in C:\\Program Files gefunden. Versuche Fallback.")
        # Fallback (falls manuell in einen anderen Ordner installiert)
        return None # pdf2image versucht dann, den PATH zu nutzen
    
    # Bei macOS/Linux wird der PATH meist korrekt von brew/apt gesetzt
    return None

# Führe die Erkennung beim Start des Skripts einmal aus
POPPLER_PATH = get_poppler_path()
# --- ENDE NEUER CODE ---


def init_paths_from_config(config):
    base_dir = os.path.dirname(__file__)
    ressources_dir = getattr(config, 'RESSOURCES_DIR', os.path.abspath(
        os.path.join(base_dir, "..", "ressources")))
    download_dir = getattr(config, 'DOWNLOAD_DIR',
                           os.path.join(ressources_dir, "downloads"))
    extract_dir = getattr(config, 'EXTRACT_DIR',
                          os.path.join(download_dir, "extracted"))
    module_map_csv = getattr(config, 'MODULE_MAP_CSV', os.path.join(
        ressources_dir, "modul_mengen_stat_vwl_bwl.csv"))
    output_csv = getattr(config, 'OUTPUT_CSV', os.path.join(
        ressources_dir, "bewerber_evaluierung.csv"))

    os.makedirs(download_dir, exist_ok=True)
    os.makedirs(extract_dir, exist_ok=True)

    return {
        "ressources_dir": ressources_dir,
        "download_dir": download_dir,
        "extract_dir": extract_dir,
        "module_map_csv": module_map_csv,
        "output_csv": output_csv,
    }

ECTS_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*CP", re.IGNORECASE)
NOTE_RE = re.compile(r"(\d(?:[.,]\d)?)")
NOTE_STRICT_RE = re.compile(r"\b([1-4][.,]\d)\b")
ROW_LOCATOR = (
    By.XPATH, "//table//tr[.//td and not(contains(@style,'display:none'))]")

def ensure_ocr_available():
    if convert_from_path is None or pytesseract is None:
        raise RuntimeError("OCR nicht verfügbar")
    return True

def set_chrome_download_dir(driver, download_dir):
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                               "behavior": "allow", "downloadPath": download_dir})
        return True
    except Exception:
        try:
            driver.execute_cdp_cmd("Browser.setDownloadBehavior", {
                                   "behavior": "allow", "downloadPath": download_dir})
            return True
        except Exception:
            return False

def wait_for_any_file(download_dir, pattern="*.zip", timeout=30, prev=None):
    prev_set = set(prev or glob.glob(os.path.join(download_dir, pattern)))
    deadline = time.time() + timeout
    while time.time() < deadline:
        current = set(glob.glob(os.path.join(download_dir, pattern)))
        new = current - prev_set
        if new:
            return sorted(list(new), key=lambda p: os.path.getmtime(p))[-1]
        time.sleep(0.5)
    return None

def extract_zip_to_dir(zip_path, target_dir):
    shutil.rmtree(target_dir, ignore_errors=True)
    os.makedirs(target_dir, exist_ok=True)
    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(target_dir)
    return target_dir

def find_pdfs_in_dir(d):
    pdfs = glob.glob(os.path.join(d, "**", "*.pdf"), recursive=True)
    return [p for p in pdfs if "__deckblatt" not in os.path.basename(p).lower() and "deckblatt" not in os.path.basename(p).lower()]

def ocr_text_from_pdf(pdf_path, dpi=200):
    if convert_from_path is None or pytesseract is None:
        raise RuntimeError("OCR nicht verfügbar")
        
    print(f"TRACE: Starte OCR für {pdf_path}")
    
    # --- START ÄNDERUNG: POPPLER-PFAD ÜBERGEBEN ---
    # Übergibt den automatisch erkannten Pfad an die Funktion
    images = convert_from_path(pdf_path, dpi=dpi, poppler_path=POPPLER_PATH)
    # --- ENDE ÄNDERUNG ---

    text_parts = []
    config = "--psm 6" 

    for img in images:
        try:
            text_parts.append(pytesseract.image_to_string(
                img, lang="deu+eng", config=config)) 
        except Exception as e:
            print(f"OCR-Fehler bei {pdf_path}: {e}")
    return "\n".join(text_parts)

def load_whitelist(csv_path):
    whitelist = set()
    if not csv_path or not os.path.exists(csv_path):
        print("Keine Whitelist-Datei angegeben.")
        return whitelist
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        next(reader, None)
        for row in reader:
            if row and row[0].strip():
                whitelist.add(row[0].strip().lower())
    print(f"Whitelist geladen: {len(whitelist)} Einträge.")
    return whitelist

def load_module_mapping(csv_path):
    mapping = {}
    if not os.path.exists(csv_path):
        print(f"Fehler: Modul-Mapping-Datei fehlt: {csv_path}")
        return mapping
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r in reader:
            key = r.get("module") or r.get("modul")
            cat = r.get("category") or r.get("Kategorie")
            if key and cat:
                mapping[key.strip().lower()] = cat.strip()
    print(f"Modul-Mapping geladen: {len(mapping)} Einträge.")
    return mapping

def is_candidate_row(row):
    try:
        cells = row.find_elements(By.TAG_NAME, "td")
        if not cells or len(cells) < 3:
            return False
        text = " ".join([c.text.strip().lower() for c in cells])
        if "bewerbung" in text or re.search(r"\b\d{5,}\b", text):
            return True
        return False
    except Exception:
        return False

def get_applicant_number_from_detail_page(browser):
    try:
        el = WebDriverWait(browser, 5).until(
            EC.presence_of_element_located(
                (By.XPATH, "//span[contains(@id, 'applicantDataSummary_number')] | //span[contains(text(), 'Bewerbernummer')]/following-sibling::span | //span[contains(text(), 'Bewerbungsnummer')]/following-sibling::span")
            )
        )
        txt = el.text.strip()
        m = re.search(r"\b(\d{5,})\b", txt)
        if m:
            return m.group(1)
        
        el_label = WebDriverWait(browser, 2).until(
            EC.presence_of_element_located(
                (By.XPATH, "//span[contains(text(), 'Bewerbernummer') or contains(text(), 'Bewerbungsnummer')]")
            )
        )
        el_value = el_label.find_element(By.XPATH, "./following-sibling::span[1]")
        txt_val = el_value.text.strip()
        m_val = re.search(r"\b(\d{5,})\b", txt_val)
        if m_val:
            return m_val.group(1)
            
        return f"unknown_{int(time.time())}"
    except Exception:
        return f"unknown_{int(time.time())}"

def clean_ocr_line(line):
    if "http" in line or ".png" in line:
        return "" 
        
    line = re.sub(r'\b[A-Z]{2,5}[ -]?\d{2,5}\b', '', line, flags=re.IGNORECASE)
    line = re.sub(r'\b[A-Z-]+\s*\d+\b', '', line)
    line = re.sub(r'\d+[.,]\d+', '', line)
    line = re.sub(r'\b\d+\b', '', line)
    line = re.sub(r'[|/\\*_]', ' ', line) 
    line = re.sub(r'Seite \d+ von \d+', '', line, flags=re.IGNORECASE)
    line = re.sub(r'Page \d+ of \d+', '', line, flags=re.IGNORECASE)
    line = re.sub(r'\b(cP|CP|Pass|Fail|Note|Grade)\b', '', line, flags=re.IGNORECASE)
    
    line = ' '.join(line.split())
    
    if len(line) < 4:
        return ""
        
    return line

def check_university_whitelist(ocr_text, whitelist_set):
    if not whitelist_set:
        return False, None
    
    text_lower = ocr_text.lower()
    for uni_name in whitelist_set:
        if uni_name in text_lower:
            return True, uni_name
    return False, None

def extract_ocr_note_from_text(text):
    possible_labels = [
        "gesamtnote", "abschlussnote", "endnote", "note", "final grade",
        "grade", "overall grade", "gesamtergebnis", "endgültige note"
    ]
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    
    for ln in lines:
        low = ln.lower()
        if any(lbl in low for lbl in possible_labels):
            m = NOTE_STRICT_RE.search(ln)
            if m:
                try:
                    val = float(m.group(1).replace(",", "."))
                    print(f"DEBUG/TRACE: OCR-Note (Label) gefunden in Zeile '{ln[:60]}...' -> {val}")
                    return val
                except ValueError:
                    continue
            
            try:
                idx = lines.index(ln)
                if idx + 1 < len(lines):
                    m2 = NOTE_STRICT_RE.search(lines[idx + 1])
                    if m2:
                        try:
                            val = float(m2.group(1).replace(",", "."))
                            print(f"DEBUG/TRACE: OCR-Note (Label) gefunden in Folgelinie -> {val}")
                            return val
                        except ValueError:
                            continue
            except Exception:
                pass 
                
    m_fallback = NOTE_STRICT_RE.search(text)
    if m_fallback:
        try:
            val = float(m_fallback.group(1).replace(",", "."))
            print(f"DEBUG/TRACE: OCR-Note (Fallback) gefunden im gesamten Text -> {val}")
            return val
        except ValueError:
            pass
            
    print("WARNUNG: Keine OCR-Note im Text gefunden.")
    return None

def match_module_in_text(mapping, text, categories):
    sums = {cat: 0.0 for cat in categories}
    matched_modules = []
    unrecognized_lines = [] 

    if not text:
        print("WARNUNG: match_module_in_text() aufgerufen mit leerem OCR-Text.")
        return sums, matched_modules, unrecognized_lines

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    FALLBACK_KEYWORDS = {
        "VWL": ["volkswirtschaft", "vwl", "mikroökonom", "makroökonom"],
        "Statistik": ["statistik", "ökonomet", "quantitative methoden"],
        "BWL": ["betriebswirtschaft", "business administration", "management", "finanz", "rechnungswesen"]
    }

    def add_unrecognized(line):
        cleaned_line = clean_ocr_line(line)
        if cleaned_line and cleaned_line not in unrecognized_lines: 
            unrecognized_lines.append(cleaned_line)

    def find_best_ects_in_vicinity(start_index):
        potential_ects = []
        best_offset = 0 # Wichtig, um processed_line_indices zu aktualisieren
        
        for i in range(min(3, len(lines) - start_index)):
            current_index = start_index + i
            line_to_check = lines[current_index]
            
            matches = ECTS_RE.findall(line_to_check)
            for m in matches:
                try:
                    val = float(m.replace(",", "."))
                    if val > 0.0 and val <= 50.0: 
                        potential_ects.append((val, i)) # Speichere (Wert, Offset)
                except ValueError:
                    continue
        
        if not potential_ects:
            return (0.0, 0)
            
        ects_over_4 = [p for p in potential_ects if p[0] > 4.0]
        if ects_over_4:
            best_pair = max(ects_over_4, key=lambda item: item[0])
            return best_pair 

        ects_whole_or_half = [p for p in potential_ects if (p[0] % 0.5) < 0.001]
        if ects_whole_or_half:
            best_pair = max(ects_whole_or_half, key=lambda item: item[0])
            return best_pair

        best_pair = max(potential_ects, key=lambda item: item[0])
        return best_pair

    processed_line_indices = set()

    for i, ln in enumerate(lines):
        low = ln.lower()
        if i in processed_line_indices:
            continue 

        matched_category = None
        module_name_found = None
        
        if mapping:
            for module_name, cat in mapping.items():
                if module_name in low:
                    if cat in sums:
                        matched_category = cat
                        module_name_found = module_name
                        break 
        
        if not matched_category:
            for cat, keywords in FALLBACK_KEYWORDS.items():
                if cat in categories and any(kw in low for kw in keywords):
                    matched_category = cat
                    module_name_found = f"Fallback: {keywords[0]}"
                    break
        
        if not matched_category:
            continue

        val, offset = find_best_ects_in_vicinity(i)

        if val > 0.0:
            sums[matched_category] += val
            matched_modules.append(f"{module_name_found}->{matched_category}:{val}")
            # Markiere alle Zeilen, die durchsucht wurden, als verbraucht
            for j in range(offset + 1):
                processed_line_indices.add(i + j)
        else:
            add_unrecognized(ln)

    sums = {k: round(v, 2) for k, v in sums.items()}
    print(f"TRACE: Modul-Matching abgeschlossen. Summen: {sums}")
    print(f"TRACE: {len(unrecognized_lines)} bereinigte, unerkannte Zeilen gefunden.")
    return sums, matched_modules, unrecognized_lines

def evaluate_requirements(claimed_note, ocr_note, ocr_data, recognized, unrecognized, config):
    reasons = []
    ok = True

    req_note_max = getattr(config, "REQ_NOTE_MAX", 2.4) 
    requirements_ects = getattr(config, "REQUIREMENTS", {})

    if ocr_note is not None:
        note_used = ocr_note
        note_source = "OCR"
    else:
        note_used = claimed_note
        note_source = "Claimed"

    if note_used is None:
        reasons.append(f"Note nicht vorhanden (Quelle: {note_source})")
        ok = False
    else:
        try:
            if note_used > req_note_max:
                reasons.append(f"Note zu schlecht ({note_used} > {req_note_max})")
                ok = False
        except Exception:
            reasons.append("Note ungültig")
            ok = False

    if ocr_note is not None and claimed_note is not None:
        diff = abs(ocr_note - claimed_note)
        if diff >= 0.1: 
            reasons.append(f"Abweichung Note (Angabe: {claimed_note}, Dokument: {ocr_note})")
            
    if not requirements_ects:
        reasons.append("Keine ECTS-Anforderungen in Config definiert.")
        ok = False
    else:
        for category, req_value in requirements_ects.items():
            ocr_value = float(ocr_data.get(category, 0.0))
            if ocr_value < req_value:
                reasons.append(f"{category} zu wenig ({ocr_value} < {req_value})")
                ok = False

    if unrecognized:
        reasons.append(f"{len(unrecognized)} unerkannte Modul(e) gefunden")

    status = "Erfuellt" if ok else "Nicht erfuellt"
    return status, "; ".join(reasons) if reasons else "Alle Kriterien erfuellt"

def extract_claimed_from_dom(browser, config):
    print("TRACE: Starte Extraktion der Claimed-Werte aus DOM...")
    
    categories = list(getattr(config, "REQUIREMENTS", {}).keys())
    
    claimed = {"note": None}
    claimed.update({cat: 0.0 for cat in categories})

    # --- START FINALE ÄNDERUNG (Problem 1: Claimed Note) ---
    # 1. VERSUCH: Robuster "Label-for -> ID"-Ansatz
    try:
        label_element = WebDriverWait(browser, 0.5).until(
            EC.presence_of_element_located((By.XPATH, "//label[normalize-space(.)='Ergebnis MZB-Note']"))
        )
        note_id = label_element.get_attribute('for')
        if note_id:
            note_element = browser.find_element(By.XPATH, f"//div[@id='{note_id}']//span[@class='applicationContentTextLineBreak']")
            text = note_element.text.strip()
            m = NOTE_STRICT_RE.search(text) or NOTE_RE.search(text)
            if m:
                claimed["note"] = float(m.group(1).replace(",", "."))
                print(f"DEBUG/TRACE: Claimed-Note (MZB) extrahiert: {claimed['note']}")
    except Exception as e:
        print(f"TRACE: 'Ergebnis MZB-Note'-Logik fehlgeschlagen ({e}). Versuche Fallbacks...")
        
    # 2. VERSUCH: Fallback-Logik (Deine "nimm die zweite Note"-Logik)
    if claimed['note'] is None:
        note_selectors = [
            "//span[contains(@id, 'applicantDataSummary_number')]/following::span[contains(@class,'applicationContentTextLineBreak')][2]",
            "//span[contains(@id, 'applicantDataSummary_number')]/following::span[contains(@class,'applicationContentTextLineBreak')][1]",
            "//span[normalize-space(.)='Abschlussnote']/following-sibling::span",
            "//span[normalize-space(.)='Gesamtnote']/following-sibling::span",
        ]
        
        for selector in note_selectors:
            try:
                el = WebDriverWait(browser, 0.5).until(
                    EC.presence_of_element_located((By.XPATH, selector)))
                text = el.text.strip()
                if text:
                    m = NOTE_STRICT_RE.search(text) or NOTE_RE.search(text)
                    if m:
                        claimed["note"] = float(m.group(1).replace(",", "."))
                        print(f"DEBUG/TRACE: Claimed-Note (Fallback) extrahiert: {claimed['note']} (mit '{selector}')")
                        break
            except Exception:
                continue
                
    if claimed['note'] is None:
        print("WARNUNG: Konnte Claimed-Note nicht aus DOM extrahieren.")
    # --- ENDE FINALE ÄNDERUNG ---


    label_xpath = "//label[contains(normalize-space(.),'CP im Bereich')]"
    dom_map = getattr(config, "DOM_ECTS_MAP", {})

    try:
        labels = browser.find_elements(By.XPATH, label_xpath)
        print(f"DEBUG/TRACE: {len(labels)} ECTS-Label ('CP im Bereich...') gefunden.")
        
        legacy_ects_re = re.compile(r"(\d+(?:[.,]\d+)?)")
        
        for lab in labels:
            txt = lab.text.strip().lower()
            cat_found = None
            
            for dom_key, mapped_cat in dom_map.items():
                if dom_key.lower() in txt:
                    if mapped_cat in categories: 
                        cat_found = mapped_cat
                        break
            
            if not cat_found:
                for cat in categories:
                    if cat.lower() in txt:
                        cat_found = cat
                        break
                        
            if not cat_found:
                continue
                
            try:
                sib = lab.find_element(By.XPATH, "following-sibling::*[1]")
                m_legacy = legacy_ects_re.search(sib.text.strip())
                
                if m_legacy:
                    val = float(m_legacy.group(1).replace(",", "."))
                    claimed[cat_found] += val
                    print(f"DEBUG/TRACE: Claimed-ECTS gefunden: {cat_found} -> {val}")
                else:
                    print(f"WARNUNG: ECTS-Wert für '{txt[:50]}...' nicht gefunden (Sibling-Text: {sib.text.strip()[:50]}).")
            except Exception:
                print(f"WARNUNG: Konnte ECTS-Sibling für '{txt[:50]}...' nicht finden/lesen.")
                continue
    except Exception as e:
        print(f"WARNUNG: Fehler bei DOM-ECTS-Auslesung: {e}")

    for k in categories:
        claimed[k] = round(claimed[k], 2)
    print(f"DEBUG/TRACE: Finale Claimed-Werte: {claimed}")
    return claimed

def run_filterphase_evaluierung(bot, flow_url, config):
    print("Starte Evaluierung...")

    paths = init_paths_from_config(config)
    try:
        ensure_ocr_available()
    except RuntimeError as e:
        print(f"FATAL: {e}. Breche Evaluierung ab.")
        return

    set_chrome_download_dir(bot.browser, paths["download_dir"])

    module_map = load_module_mapping(paths["module_map_csv"])
    whitelist_set = load_whitelist(getattr(config, "WHITELIST_UNIS", None))
    
    categories = list(getattr(config, "REQUIREMENTS", {}).keys())
    if not categories:
        print("FATAL: Keine 'REQUIREMENTS' im Config-Objekt gefunden. Breche ab.")
        return

    with open(paths["output_csv"], "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        header = ["Bewerbernummer", "Claimed_Note", "OCR_Note"]
        header.extend([f"OCR_{c}" for c in categories])
        header.extend([f"Claimed_{c}" for c in categories])
        header.extend(["Status", "Details", "MatchedModules",
                      "UnrecognizedLines"]) 
        writer.writerow(header)

    try:
        print("suche .....")
        search_btn = WebDriverWait(bot.browser, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[.//span[normalize-space()='Suchen']] | //span[normalize-space()='Suchen']/parent::button | //button[contains(@id,'search')]"))
        )
        bot.browser.execute_script(
            "arguments[0].scrollIntoView(true);", search_btn)
        time.sleep(0.5)
        bot.browser.execute_script("arguments[0].click();", search_btn)
        print(" Warte auf Ergebnisse...")
        WebDriverWait(bot.browser, 15).until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "span.dataScrollerResultText")))
        print("DEBUG: Suchergebnisse geladen.")
    except Exception as e:
        print(f"FEHLER: Ergebnisse nicht laden. Breche ab. Fehler: {e}")
        return

    total_from_scroller = 0
    total = 0
    try:
        res_text_element = WebDriverWait(bot.browser, 5).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "span.dataScrollerResultText")))
        res_txt = res_text_element.text
        m = re.search(r"(\d+)", res_txt)
        if m:
            total_from_scroller = int(m.group(1))
            total = total_from_scroller
            print(
                f"TRACE Gesamtzahl der Bewerber laut dataScrollerResultText: {total}")
        else:
            raise ValueError("Zahl nicht gefunden")
    except Exception:
        print("WARNUNG: Konnte dataScrollerResultText nicht lesen. Fahre mit Zählung fort.")
        
    try:
        rows_initial = bot.browser.find_elements(*ROW_LOCATOR)
        candidate_rows_count = len(
            [r for r in rows_initial if is_candidate_row(r)])
        
        if total != candidate_rows_count:
            print(f"WARNUNG: Diskrepanz! DataScroller meldet {total_from_scroller}, aber {candidate_rows_count} reale Zeilen gefunden. Vertraue Zählung.")
            total = candidate_rows_count 
        else:
             print(f"TRACE: Zählung der Zeilen ({candidate_rows_count}) stimmt mit Scroller überein.")

    except Exception as count_e:
        if total == 0: 
            print(
                f"FEHLER: Konnte Zeilen nicht finden/zählen. Breche ab. Fehler: {count_e}")
            return

    if total == 0:
        print("keine bwerebr gefunde")
        return

    main_window_handle = bot.browser.current_window_handle
    print(f"DEBUG: Handle: {main_window_handle}")

    for i in range(total):
        new_tab_handle = None
        applicant_num_from_list = f"unknown_idx_{i}"
        applicant_num = applicant_num_from_list
        ocr_note = None 
        saved_pdf_counts = {cat: 0.0 for cat in categories} 
        matched_modules = []
        unrecognized_lines = []
        full_ocr_text = ""
        priority_grade_text = "" 

        try:
            print(f"--- Verarbeitung von Bewerber {i+1}/{total} (Index {i}) ---")
            if bot.browser.current_window_handle != main_window_handle:
                bot.browser.switch_to.window(main_window_handle)
            time.sleep(0.5)

            rows = WebDriverWait(bot.browser, 10).until(
                EC.presence_of_all_elements_located(ROW_LOCATOR))
            candidate_rows = [r for r in rows if is_candidate_row(r)]

            if i >= len(candidate_rows):
                print(
                    f"WARNUNG: Index {i} außerhalb ({len(candidate_rows)}). Breche Schleife ab.")
                break
            current_row = candidate_rows[i] 

            try:
                td_num = current_row.find_element(
                    By.XPATH, ".//td[contains(@class,'column3') or contains(@class,'column 3')][1]")
                row_text = td_num.text.strip()
                mnum = re.search(r"\b(\d{5,})\b", row_text)
                if mnum:
                    applicant_num_from_list = mnum.group(1)
                    applicant_num = applicant_num_from_list
            except Exception:
                pass

            url_to_open = None
            element_for_js_click = None
            try:
                link_element = current_row.find_element(
                    By.XPATH, ".//a[contains(@href,'applicationEditor-flow')]")
                url_to_open = link_element.get_attribute('href')
                element_for_js_click = link_element
            except NoSuchElementException:
                try:
                    button_element = current_row.find_element(
                        By.XPATH, ".//button[contains(@id,'tableRowAction') or contains(@name,'tableRowAction')]")
                    element_for_js_click = button_element
                except NoSuchElementException:
                    print(
                        f"WARNUNG: Weder Link noch Button für Bewerber {applicant_num_from_list} gefunden. Überspringe.")
                    continue

            print(
                f"DEBUG: Öffne Bewerber {applicant_num_from_list} ({i+1}) in neuem Tab...")
            initial_handles = set(bot.browser.window_handles)

            if url_to_open:
                bot.browser.execute_script(
                    f"window.open('{url_to_open}', '_blank');")
            elif element_for_js_click:
                bot.browser.execute_script(
                    "arguments[0].click();", element_for_js_click)
            else:
                print("FEHLER: Kein Element zum Klicken. Überspringe.")
                continue

            time.sleep(3) 
            current_handles = bot.browser.window_handles
            new_handles = set(current_handles) - initial_handles
            
            if not new_handles and "applicationEditor-flow" not in bot.browser.current_url:
                print(f"FEHLER: Neuer Tab nicht geöffnet (oder falsche URL). Überspringe.")
                try: bot.browser.switch_to.alert.dismiss()
                except: pass
                continue
            elif not new_handles and "applicationEditor-flow" in bot.browser.current_url:
                print("DEBUG: Navigation im selben Tab erkannt.")
                new_tab_handle = main_window_handle
            else:
                new_tab_handle = list(new_handles)[0]
                bot.browser.switch_to.window(new_tab_handle)
                print(f"DEBUG: Zum neuen Tab gewechselt: {new_tab_handle}")

            if i == 0:
                print(
                    "INFO: Erster Bewerber (Index 0) — warte bis zu 7 Sekunden, damit Popup manuell geschlossen werden kann...")
                try:
                    # 2  (7 mit 0.5 ersetzen)
                    WebDriverWait(bot.browser, 7).until(
                        EC.any_of(
                            EC.presence_of_element_located(
                                (By.XPATH, "//span[contains(text(),'Bewerbernummer')]")),
                            EC.presence_of_element_located(
                                (By.XPATH, "//button[contains(@id,'showRequestSubjectBtn')]")),
                            EC.presence_of_element_located(
                                (By.XPATH, "//label[contains(normalize-space(.),'CP im Bereich')]"))
                        )
                    )
                    time.sleep(0.5)
                except Exception:
                    print("INFO: Timeout bei Warten auf Element. Gehe von manuellem Popup aus.")
                    time.sleep(7) 
                print(
                    "INFO: Popup-Wartezeit vorbei — setze automatische Verarbeitung fort...")

            WebDriverWait(bot.browser, 15).until(lambda d: d.execute_script("return document.readyState") == "complete")
            time.sleep(1) 

            applicant_num = get_applicant_number_from_detail_page(bot.browser)
            print(
                f"DEBUG: Aktuelle Bewerbernummer im Detail-Tab: {applicant_num}")

            try:
                application_buttons = bot.browser.find_elements(
                    By.XPATH, "//button[contains(@id, 'showRequestSubjectBtn')]")
                if application_buttons:
                    btn_text = application_buttons[0].text
                    print(
                        f"INFO: {len(application_buttons)} Anträge. Wähle ersten: '{btn_text}'")
                    bot.browser.execute_script(
                        "arguments[0].click();", application_buttons[0])
                    WebDriverWait(bot.browser, 10).until(lambda d: d.execute_script(
                        "return document.readyState") == "complete")
                    time.sleep(2)
            except Exception as e:
                print(
                    f"DEBUG: Keine separaten Antrags-Buttons oder Fehler: {e}")

            dl_element = None
            for xp in [
                "//button[contains(@aria-label,'Nachweise herunterladen') or contains(@title,'Nachweise herunterladen')]",
                "//button[.//img[contains(@src,'download.svg') or @alt='Nachweise herunterladen']]",
                "//img[@alt='Nachweise herunterladen']/ancestor::button[1]",
                "//a[.//img[contains(@src,'download.svg')]]",
                "//a[contains(text(), 'Download') or contains(text(), 'ZIP')]"
            ]:
                try:
                    dl_element = WebDriverWait(bot.browser, 3).until(
                        EC.presence_of_element_located((By.XPATH, xp)))
                    if dl_element:
                        print(f"TRACE Download-Element gefunden mit XPath: {xp}")
                        break
                except Exception:
                    dl_element = None

            for f in glob.glob(os.path.join(paths["download_dir"], "*")):
                try:
                    if os.path.isfile(f):
                        os.remove(f)
                    else:
                        shutil.rmtree(f, ignore_errors=True)
                except Exception:
                    pass
            os.makedirs(paths["extract_dir"], exist_ok=True)

            if dl_element:
                print("/TRACE:starte Download...")
                prev_zips = glob.glob(os.path.join(
                    paths["download_dir"], "*.zip"))
                try:
                    bot.browser.execute_script(
                        "arguments[0].click();", dl_element)
                except Exception as e:
                    print(f"Download-Klick fehlgeschlagen: {e}")

                print("Warte auf ZIP-Datei...")
                zip_found = wait_for_any_file(
                    paths["download_dir"], pattern="*.zip", timeout=35, prev=prev_zips)
                if not zip_found:
                    print(f"Kein ZIP-Download für {applicant_num}.")
                else:
                    print(f" ZIP gefunden: {zip_found}. Entpacke...")
                    try:
                        extract_dir = os.path.join(
                            paths["extract_dir"], f"{applicant_num}_{int(time.time())}")
                        extract_zip_to_dir(zip_found, extract_dir)
                        print(
                            f"TRACE: Entpackt nach: {extract_dir}. Suche PDFs...")
                        pdfs = find_pdfs_in_dir(extract_dir)
                        print(
                            f"DEBUG/TRACE: {len(pdfs)} PDFs gefunden (Deckblätter entfernt): {[os.path.basename(p) for p in pdfs]}")
                        
                        if not pdfs:
                            print(
                                f"WARNUNG: Keine relevanten PDFs in ZIP für {applicant_num}.")
                        else:
                            grade_keywords = ["zeugnis", "vpd", "certificate", "urkunde", "bachelor-zeugnis"]
                            
                            priority_grade_pdfs = [p for p in pdfs if any(kw in os.path.basename(p).lower() for kw in grade_keywords)]

                            for pdf_path in pdfs:
                                try:
                                    txt = ocr_text_from_pdf(pdf_path)
                                    full_ocr_text += "\n" + txt
                                    if pdf_path in priority_grade_pdfs:
                                        priority_grade_text += "\n" + txt
                                except Exception as e:
                                    print(
                                        f"WARNUNG: OCR Fehler bei {pdf_path}: {e}")
                            
                            print(
                                f"DEBUG/TRACE: Gesamter OCR-Text (erste 500 Zeichen):\n{full_ocr_text[:500]}...")

                            is_whitelisted, uni_match = check_university_whitelist(
                                full_ocr_text, whitelist_set)

                            if is_whitelisted:
                                print(
                                    f"INFO: Bewerber {applicant_num} ist auf der Whitelist (Match: '{uni_match}').")
                                status = "Zugelassen (Whitelist)"
                                details = f"Uni-Whitelist: {uni_match}"

                                ocr_note = extract_ocr_note_from_text(priority_grade_text or full_ocr_text)
                                claimed = extract_claimed_from_dom(
                                    bot.browser, config)

                                csv_row = [
                                    applicant_num,
                                    claimed.get("note"),
                                    ocr_note,
                                ]
                                for cat in categories:
                                    csv_row.append(0.0)  
                                for cat in categories:
                                    csv_row.append(claimed.get(cat, 0.0)) 
                                
                                csv_row.extend(
                                    [status, details, "N/A (Whitelist)", "N/A (Whitelist)"]) 

                                with open(paths["output_csv"], "a", newline="", encoding="utf-8") as of:
                                    writer = csv.writer(of)
                                    writer.writerow(csv_row)

                                print(
                                    f"DEBUG: Ergebnis für {applicant_num} (Whitelist) geschrieben.")
                                continue 

                            print(
                                f"INFO: Bewerber {applicant_num} nicht auf Whitelist. Starte ECTS/Noten-Prüfung...")

                            ocr_note = None
                            if priority_grade_text:
                                print("DEBUG/TRACE: Suche Note in Prioritäts-Dokumenten (Zeugnis, VPD)...")
                                ocr_note = extract_ocr_note_from_text(priority_grade_text)
                            
                            if ocr_note is None:
                                print("DEBUG/TRACE: Keine Note in Prioritäts-Dokumenten gefunden. Suche im gesamten Text (Fallback)...")
                                ocr_note = extract_ocr_note_from_text(full_ocr_text)

                            sums, matched_modules, unrecognized_lines = match_module_in_text(
                                module_map, full_ocr_text, categories)

                            saved_pdf_counts = sums

                    except Exception as e:
                        print(
                            f"WARNUNG: Fehler beim Entpacken/OCR für {applicant_num}: {e}")
                        saved_pdf_counts = {cat: 0.0 for cat in categories}
                        matched_modules = []
                        unrecognized_lines = []
            else:
                print(
                    f"WARNUNG: Kein Download-Element für {applicant_num} (kein ZIP).")
                ocr_note = None 

            claimed = extract_claimed_from_dom(bot.browser, config)

            status, details = evaluate_requirements(
                claimed.get("note"),
                ocr_note, 
                saved_pdf_counts, 
                matched_modules,
                unrecognized_lines, 
                config
            )
            
            csv_row = [
                applicant_num,
                claimed.get("note"),
                ocr_note
            ]
            for cat in categories:
                csv_row.append(saved_pdf_counts.get(cat, 0.0))
            for cat in categories:
                csv_row.append(claimed.get(cat, 0.0))
                
            csv_row.extend([
                status,
                details,
                " | ".join(matched_modules),
                " | ".join(unrecognized_lines) 
            ])

            with open(paths["output_csv"], "a", newline="", encoding="utf-8") as of:
                writer = csv.writer(of)
                writer.writerow(csv_row)

            print(
                f"DEBUG: Ergebnis für {applicant_num} geschrieben. Status: {status}. Details: {details}")

        except Exception as e:
            print(
                f"FATALER FEHLER bei der Verarbeitung von Bewerber {i+1} ({applicant_num_from_list}): {e}")
            try:
                with open(paths["output_csv"], "a", newline="", encoding="utf-8") as of:
                    writer = csv.writer(of)
                    error_row = [applicant_num] + ["ERROR"] * (len(categories) * 2 + 2) + ["FATAL ERROR", str(e), "", ""] 
                    writer.writerow(error_row)
            except Exception as write_e:
                print(f"Konnte FATAL ERROR nicht in CSV schreiben: {write_e}")


        finally:
            # Tab schließen
            try:
                current_handle = bot.browser.current_window_handle
                if current_handle != main_window_handle and current_handle in bot.browser.window_handles:
                    print(f"DEBUG: Schließe Tab {current_handle}...")
                    bot.browser.close()

                time.sleep(0.5)

                if main_window_handle in bot.browser.window_handles:
                    bot.browser.switch_to.window(main_window_handle)
                    print(f"DEBUG: back zum Haupt-Tab {main_window_handle}.")
                else:
                    remaining_handles = bot.browser.window_handles
                    if len(remaining_handles) == 1:
                        main_window_handle = remaining_handles[0] # Handle "retten"
                        bot.browser.switch_to.window(main_window_handle)
                        print(
                            f"WARNUNG: Haupt-Tab-Handle gerettet. Back zum Haupt-Tab {main_window_handle}.")
                    else:
                        print(
                            "FATAL: Haupt-Tab wurde unerwartet geschlossen. Breche ab.")
                        raise RuntimeError("Haupt-Tab verloren")
            except Exception as e_finally:
                print(
                    f"FATAL: Kritischer Fehler im 'finally'-Block: {e_finally}")
                raise RuntimeError("Fehler im 'finally'-Block")

    print("DEBUG: Evaluierungs-Phase abgeschlossen. CSV:", paths["output_csv"])