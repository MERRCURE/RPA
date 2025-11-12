import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class BrowserAutomation:
    def __init__(self, options=None):
        if options:
            self.browser = webdriver.Chrome(options=options)
        else:
            self.browser = webdriver.Chrome()
        self.wait = WebDriverWait(self.browser, 2)

    def open_url(self, url: str):
        self.browser.get(url)

        try:
            WebDriverWait(self.browser, 15).until(
                lambda d: d.execute_script(
                    "return document.readyState") == "complete"
            )
        except Exception:
            print("Timeout: kein readyState=complete")

        try:
            body_visible = WebDriverWait(self.browser, 5).until(
                lambda d: d.execute_script(
                    "return document.body.style.visibility") != "hidden"
            )
            if body_visible:
                print("STATUS: Seite sichtbar.")
        except Exception:
            self.browser.execute_script(
                "document.body.style.visibility='visible';")
            print(" sichtbarkeit manuell ")

    def add_input(self, by: By, value: str, text: str):
        """auf DOM praesenz warten"""
        try:
            field = WebDriverWait(self.browser, 15).until(
                EC.presence_of_element_located((by, value))
            )
            time.sleep(0.5)
            self.browser.execute_script(
                "arguments[0].scrollIntoView(true);", field)
            field.clear()
            field.send_keys(text)
            print(f"DEBUG: Text '{text}' in Feld {value} eingetragen.")
        except Exception as e:
            print(f"ERROR: Konnte Feld {value} nicht ausfüllen: {e}")

    def click_button(self, by: By, value: str):
        try:
            button = WebDriverWait(self.browser, 10).until(
                EC.element_to_be_clickable((by, value))
            )
            button.click()
            print(f"STATUS: Button {value} geklickt.")
        except Exception as e:
            print(f"ERROR: Button {value} nicht klickbar: {e}")
