import os

from PyQt6.QtCore import QSettings
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
import re

from selenium import webdriver
from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import WebDriverWait
from time import sleep

from webdriver_manager.chrome import ChromeDriverManager


class Whatsapp:

    driver = None
    wait = None

    def __init__(self, settings: QSettings):
        self.settings = settings
        self.options = webdriver.ChromeOptions()
        self.options.add_argument('--allow-profiles-outside-user-dir')
        self.options.add_argument('--enable-profile-shortcut-manager')
        self.options.add_argument(r'user-data-dir=' + settings.value("whatsapp/cache_dir"))
        self.options.add_argument('--profile-directory=Hacker')
        self.options.add_argument('--profiling-flush=n')
        self.options.add_argument('--enable-aggressive-domstorage-flushing')
        self.options.add_argument('--disable-dev-shm-usage')
        self.options.add_argument('--no-sandbox')

    def authorization(self):
        try:
            if self.driver is None:
                self.driver =  webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=self.options)
                self.wait = WebDriverWait(self.driver, 30)

            self.driver.get('https://web.whatsapp.com')
            if len(os.listdir(self.settings.value("whatsapp/cache_dir"))) == 0:
                sleep(60)
        except Exception as e:
            print(e)
            return False

        return True

    def send_spam(self, phone, text):
        clear_phone = re.sub('[() -]', '', phone)

        try:
            url = f"https://web.whatsapp.com/send?phone={clear_phone}&text={text.replace(' ', '+').replace('\n', '%0a')}"
            self.driver.get(url)
            send_button_xpath = self.settings.value("whatsapp/send_button_xpath")
            self.wait.until(EC.element_to_be_clickable((By.XPATH, send_button_xpath)))
            self.driver.find_element(By.XPATH, send_button_xpath).click()
            self.wait.until(EC.element_to_be_clickable((By.XPATH, send_button_xpath)))
        except Exception as e:
            return "[" + str(clear_phone) + "]: " + "Ошибка отправки сообщения " + str(e)

        return "[" + str(clear_phone) + "]: " + str(text)

    def close_browser(self):
        try:
            sleep(5) #подождать, чтобы последнее сообщение точно отправилось
            self.driver.quit()
        except Exception as e:
            print(e)