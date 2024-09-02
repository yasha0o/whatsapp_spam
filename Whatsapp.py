import logging
from time import sleep

from PySide6.QtCore import QSettings
from selenium import webdriver
from selenium.common import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


class Whatsapp:

    driver = None
    wait = None

    def __init__(self, settings: QSettings):
        self.settings = settings

        self.options = webdriver.ChromeOptions()
        self.options.add_argument('--allow-profiles-outside-user-dir')
        self.options.add_argument('--enable-profile-shortcut-manager')
        self.options.add_argument(r'user-data-dir=' + str(settings.value("whatsapp/cache_dir")))
        self.options.add_argument('--profile-directory=Hacker')
        self.options.add_argument('--profiling-flush=n')
        self.options.add_argument('--enable-aggressive-domstorage-flushing')
        self.options.add_argument('--disable-dev-shm-usage')
        self.options.add_argument('--no-sandbox')

        if str(settings.value("whatsapp/headless")).lower() == 'true':
            self.options.add_argument('--headless')

    def lazy_init(self):
        logging.info("lazy_init")
        if self.driver is None:
            logging.info("self.driver is None")
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=self.options)
            self.wait = WebDriverWait(self.driver, float(str(self.settings.value("whatsapp/timeout"))))

    def authorization(self):
        logging.info("authorization")
        driver = None
        wait = None

        try:
            auth_options = webdriver.ChromeOptions()
            auth_options.add_argument('--allow-profiles-outside-user-dir')
            auth_options.add_argument('--enable-profile-shortcut-manager')
            auth_options.add_argument(r'user-data-dir=' + str(self.settings.value("whatsapp/cache_dir")))
            auth_options.add_argument('--profile-directory=Hacker')
            auth_options.add_argument('--profiling-flush=n')
            auth_options.add_argument('--enable-aggressive-domstorage-flushing')
            auth_options.add_argument('--disable-dev-shm-usage')
            auth_options.add_argument('--no-sandbox')

            logging.info("create driver")
            driver =  webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=auth_options)
            wait = WebDriverWait(driver, float(str(self.settings.value("whatsapp/timeout"))))
            logging.info("load https://web.whatsapp.com")
            driver.get('https://web.whatsapp.com')
            qr_path = str(self.settings.value('whatsapp/qr_path'))

            logging.info("check_auth check chats")
            if not self.check_auth(wait):
                logging.info("check_auth hasn't chats. check qr....")
                wait.until(EC.visibility_of_element_located((By.XPATH, qr_path)))
            else:
                logging.info("check_auth has chats")
                return True

            logging.info("check_auth check qr in while")
            while driver.find_element(By.XPATH, qr_path).is_displayed():
                logging.info("waiting authorization....")
                sleep(5)

            logging.info("close driver")
            driver.close()
        except NoSuchElementException as ne:
            logging.error("NoSuchElementException", exc_info=ne)
            logging.info("check auth")
            if wait is None or not self.check_auth(wait):
                logging.info("check_auth false")
                return False

            if driver is not None:
                logging.info("close driver")
                driver.close()
            return True
        except Exception as ex:
            logging.error("Exception",exc_info=ex)
            return False
        return True

    def check_auth(self, wait):
        logging.info("check_auth")
        chats_path = str(self.settings.value('whatsapp/chats_path'))
        try:
            wait.until(EC.visibility_of_element_located((By.XPATH, chats_path)))
            return True
        except TimeoutException as te:
            logging.error("TimeoutException",exc_info=te)
            return False


    def send_spam(self, phone, text):
        logging.info(f"send_spam phone = {phone}, text = {text}")
        self.lazy_init()

        logging.info(f"clear phone = {phone}")
        try:
            url = f"https://web.whatsapp.com/send?phone={phone}&text={text.replace(' ', '+').replace('\n', '%0a')}"
            logging.info(f"get url = {url}")
            self.driver.get(url)
            send_button_xpath = str(self.settings.value("whatsapp/send_button_xpath"))
            logging.info("wait send button")
            self.wait.until(EC.element_to_be_clickable((By.XPATH, send_button_xpath)))
            logging.info("click send button")
            self.driver.find_element(By.XPATH, send_button_xpath).click()
        except TimeoutException as te:
            logging.error(f"phone = {phone} TimeoutException",exc_info=te)
            return "[" + str(phone) + "]: " + "Таймаут отправки сообщения, скорее всего у абонента нет whatsapp"
        except Exception as ex:
            logging.error(f"phone = {phone} Exception",exc_info=ex)
            return "[" + str(phone) + "]: " + "Ошибка отправки сообщения" + str(ex)

        logging.info(f"send success phone = {phone}")
        return "[" + str(phone) + "]: " + str(text)

    def close_browser(self):
        logging.info("close_browser")
        try:
            self.driver.close()
            self.driver = None
            self.wait = None
        except Exception as ex:
            logging.error("Exception",exc_info=ex)