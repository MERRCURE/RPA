import time
import json
import os
import sys
import argparse
import importlib
import sys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from utils.browserautomation import BrowserAutomation
from phases.filterphase_evaluierung import run_filterphase_evaluierung


#1 URL
FLOW_URL = "https://test02.digstu.hhu.de/qisserver/pages/startFlow.xhtml?_flowId=searchApplicants-flow&navigationPosition=hisinoneapp,applicationEditorGeneratedJSFDtos&recordRequest=true"




def create_chrome_options():
    chrome_options = Options()
    
 
    chrome_options.add_argument("--window-size=1400,900") 
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--enable-javascript")

    chrome_options.add_experimental_option("prefs", {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False
    })

    return chrome_options


def perform_login(bot, username, password):
    print("STATUS: Loginformular wird ausgefüllt...")

    try:
        wait = WebDriverWait(bot.browser, 15)
        user_field = wait.until(
            EC.presence_of_element_located((By.ID, "asdf")))
        pass_field = wait.until(
            EC.presence_of_element_located((By.ID, "fdsa")))
        login_btn = wait.until(
            EC.element_to_be_clickable((By.ID, "loginForm:login")))

        user_field.clear()
        user_field.send_keys(username)
        pass_field.clear()
        pass_field.send_keys(password)
        login_btn.click()

        print("DEBUG: Login-Felder ausgefüllt und Button geklickt.")
    except Exception as e:
        print(f"FEHLER: Formular-Login fehlgeschlagen: {e}")
        return False

    try:

        WebDriverWait(bot.browser, 15).until(
            lambda d: "startFlow" in d.current_url or "portal" in d.current_url
        )
        print("STATUS: Login erfolgreich und Weiterleitung erkannt.")
        return True
    except Exception:
        print("WARNUNG: Keine automatische Weiterleitung – eventuell Popup aktiv.")
        return False


def open_flow(bot):
    print("STATUS: Öffne Flow-Seite...")
    bot.open_url(FLOW_URL)
    WebDriverWait(bot.browser, 15).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    print("STATUS: Flow-Seite geladen.")


if __name__ == "__main__":

    parser = argparse.ArgumentParser(
        description="Startet die Bewerber-Evaluierung.")
    parser.add_argument(
        "-c", "--config",
        help="Name der zu ladenden Konfigurationsdatei (z.B. 'bwl_master_config')",
        required=True
    )
    args = parser.parse_args()

    config_name = args.config
    try:

        config_module = importlib.import_module(f"config.{config_name}")
        print(f"INFO: Konfiguration '{config_name}' erfolgreich geladen.")

    except ImportError:
        print(
            f"FATAL: Konfigurations-Modul '{config_name}.py' nicht gefunden.")
        sys.exit(1)

    args = parser.parse_args()
    credentials_path = os.path.join(
        os.path.dirname(__file__), "credentials.json")
    with open(credentials_path, "r", encoding="utf-8") as f: 
        credentials = json.load(f)

    username = credentials["username"]
    password = credentials["password"]

    chrome_options = create_chrome_options()
    bot = BrowserAutomation(options=chrome_options)


    # 1 url
    login_url = "https://test02.digstu.hhu.de/qisserver/pages/cs/sys/portal/hisinoneStartPage.faces"
    print("STATUS: Öffne Login-Seite...")
    bot.open_url(login_url)
    print("STATUS: Seite sichtbar und bereit.")

    perform_login(bot, username, password)

    print("PAUSE: Falls Popup erscheint, bitte Benutzername/Passwort eingeben (7 Sekunden)...")
    #2 entfernen
    time.sleep(7)

    open_flow(bot)
    run_filterphase_evaluierung(bot, FLOW_URL, config_module)

    print("STATUS: DONE")
    input("ENTER = finish ")
