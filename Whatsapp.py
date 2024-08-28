import os

from PyQt6.QtCore import QSettings
from selenium.common import TimeoutException, NoSuchElementException
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
        self.options.add_argument('--headless')

    def lazy_init(self):
        if self.driver is None:
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=self.options)
            self.wait = WebDriverWait(self.driver, self.settings.value("whatsapp/timeout"))

    def authorization(self):
        driver = None
        wait = None

        try:
            auth_options = webdriver.ChromeOptions()
            auth_options.add_argument('--allow-profiles-outside-user-dir')
            auth_options.add_argument('--enable-profile-shortcut-manager')
            auth_options.add_argument(r'user-data-dir=' + self.settings.value("whatsapp/cache_dir"))
            auth_options.add_argument('--profile-directory=Hacker')
            auth_options.add_argument('--profiling-flush=n')
            auth_options.add_argument('--enable-aggressive-domstorage-flushing')
            auth_options.add_argument('--disable-dev-shm-usage')
            auth_options.add_argument('--no-sandbox')

            driver =  webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=auth_options)
            wait = WebDriverWait(driver, self.settings.value("whatsapp/timeout"))
            driver.get('https://web.whatsapp.com')
            qr_path = self.settings.value('whatsapp/qr_path')

            if not self.check_auth(wait):
                print("hasn't chats. check qr....")
                wait.until(EC.visibility_of_element_located((By.XPATH, qr_path)))
            else:
                return True

            while driver.find_element(By.XPATH, qr_path).is_displayed():
                print("waiting authorization....")
                sleep(5)
            driver.close()
        except NoSuchElementException:
            if wait is None or not self.check_auth(wait):
                return False
            if driver is not None:
                driver.close()
            return True
        except Exception as e:
            print(e)
            return False
        return True

    def check_auth(self, wait):
        chats_path = self.settings.value('whatsapp/chats_path')
        try:
            wait.until(EC.visibility_of_element_located((By.XPATH, chats_path)))
            return True
        except TimeoutException:
            return False


    def send_spam(self, phone, text):
        self.lazy_init()

        clear_phone = re.sub('[() -]', '', phone)

        try:
            url = f"https://web.whatsapp.com/send?phone={clear_phone}&text={text.replace(' ', '+').replace('\n', '%0a')}"
            self.driver.get(url)
            send_button_xpath = self.settings.value("whatsapp/send_button_xpath")
            self.wait.until(EC.element_to_be_clickable((By.XPATH, send_button_xpath)))
            self.driver.find_element(By.XPATH, send_button_xpath).click()
        except TimeoutException:
            return "[" + str(clear_phone) + "]: " + "Таймаут отправки сообщения, скорее всего у абонента нет whatsapp"
        except Exception as e:
            return "[" + str(clear_phone) + "]: " + "Ошибка отправки сообщения" + str(e)

        return "[" + str(clear_phone) + "]: " + str(text)

    def close_browser(self):
        try:
            sleep(5) #подождать, чтобы последнее сообщение точно отправилось
            self.driver.close()
            self.driver = None
            self.wait = None
        except Exception as e:
            print(e)