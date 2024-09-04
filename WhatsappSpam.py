import logging
import sys
import threading
from array import array
from datetime import datetime
from time import sleep

from PySide6.QtCore import Signal, Slot, QSettings, QSize
from PySide6.QtGui import QIcon,QMovie
from PySide6.QtWidgets import (
    QApplication,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QPlainTextEdit,
    QSplitter,
    QMainWindow, QLabel, QSpinBox, QLineEdit, QFileDialog, QMessageBox, QStyle, QDialog, QWidget
)

import Whatsapp
from ExcelReader import ExcelReader


def color_mode(mode):
    if mode == "test":
        return '[<span style="color: green;">' + str(mode) + '</span>]'
    if mode == "whatsapp":
        return '[<span style="color: red;">' + str(mode) + '</span>]'
    if mode == "excel":
        return '[<span style="color: blue;">' + str(mode) + '</span>]'

    return '[' + str(mode) + ']'


class WhatsappSpamWindow(QWidget):
    authorization_end = Signal(bool)
    spam_sent = Signal(str)
    log_signal = Signal(str, str)
    settings = QSettings("settings.ini", QSettings.Format.IniFormat)

    whatsapp = Whatsapp.Whatsapp(settings)
    excel_reader = ExcelReader()

    def __init__(self):
        super().__init__()

        logging.basicConfig(filename=str(self.settings.value('logs/log_path')),
                            filemode='w',
                            format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s',
                            datefmt='%H:%M:%S',
                            level=logging.DEBUG)
        logging.info(f"start settings = {self.settings.allKeys()}")

        self.phones = []
        self.excel_reader.end_signal.connect(self.read_excel_end)

        self.setObjectName("window")
        self.setStyleSheet("QWidget#window { background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1, "
                           "stop: 0 #ededed, stop: 1 #e0e2e5);}"
                           "QPushButton {background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1, "
                           "stop: 0 #dadae2, stop: 1 #e3e5e7);"
                           "padding: 2px;"
                           "padding-left: 6px; "
                           "padding-right: 6px; }")

        self.setWindowTitle("Рассылка сообщений в whatsapp")
        self.setWindowIcon(QIcon("resources/icon.png"))
        layout = QVBoxLayout()

        self.file_line = QLineEdit()
        self.file_line.setReadOnly(True)
        self.file_line.textChanged.connect(self.read_excel)
        self.file_button = QPushButton("Выбрать файл...")
        self.file_button.clicked.connect(self.on_file_button_clicked)

        file_layout = QHBoxLayout()
        file_layout.addWidget(self.file_line)
        file_layout.addWidget(self.file_button)

        layout.addLayout(file_layout)

        self.phone_count_label = QLabel()

        load_layout = QGridLayout()

        download_layout = QHBoxLayout()
        download_layout.addWidget(QLabel("Загружено телефонов для рассылки:"))
        download_layout.addWidget(self.phone_count_label)
        download_layout.addWidget(QSplitter())
        load_layout.addLayout(download_layout, 0, 0, 1, 1)

        text_label_layout = QHBoxLayout()
        text_label_layout.addWidget(QLabel("Текст рассылки:"))
        text_label_layout.addWidget(QSplitter())
        load_layout.addLayout(text_label_layout, 1, 0, 1, 1)

        self.load_gif = QLabel()
        self.load_gif.setFixedSize(QSize(105, 60))
        self.load_gif.setVisible(False)
        self.gif = QMovie("resources/load.gif")
        self.gif.setScaledSize(self.load_gif.size())
        self.load_gif.setMovie(self.gif)
        load_layout.addWidget(self.load_gif, 0, 1, 2, 1)
        load_layout.setRowMinimumHeight(0, 30)
        load_layout.setRowMinimumHeight(1, 30)

        layout.addLayout(load_layout)

        self.spam_text = QPlainTextEdit()
        text_layout = QVBoxLayout()
        text_layout.addWidget(self.spam_text)

        layout.addLayout(text_layout)

        self.send_delay = QSpinBox()
        self.send_delay.setValue(5)
        self.send_delay.setMinimumSize(QSize(40, 10))

        self.test_button = QPushButton("Тест")
        self.test_button.clicked.connect(self.on_test_button_clicked)
        self.send_button = QPushButton("Отправить")
        self.send_button.clicked.connect(self.on_send_button_clicked)
        self.send_button.setEnabled(False)

        self.spam_sent.connect(self.send_whatsapp_end)

        send_layout = QHBoxLayout()
        send_layout.addWidget(QLabel("Задержка отправки:"))
        send_layout.addWidget(self.send_delay)
        send_layout.addWidget(QLabel("сек."))
        send_layout.addWidget(QSplitter())
        send_layout.addWidget(self.test_button)
        send_layout.addWidget(self.send_button)

        layout.addLayout(send_layout)

        self.authorization_button = QPushButton("Проверить авторизацию / Войти")
        self.authorization_button.clicked.connect(self.on_authorization_button_clicked)
        self.authorization_label = QLabel("<-- Для разблокировки отправки авторизуйтесь")
        self.authorization_end.connect(self.authorization_whatsapp_end)

        authorization_layout = QHBoxLayout()
        authorization_layout.addWidget(self.authorization_button)
        authorization_layout.addWidget(QSplitter())
        authorization_layout.addWidget(self.authorization_label)

        layout.addLayout(authorization_layout)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log_signal.connect(self.logging)
        self.excel_reader.log_signal.connect(self.logging)

        log_layout = QVBoxLayout()
        log_layout.addWidget(QLabel("Логирование"))
        log_layout.addWidget(self.log)

        layout.addLayout(log_layout)

        self.setLayout(layout)

    def on_authorization_button_clicked(self):
        logging.info("on_authorization_button_clicked")

        self.logging("whatsapp", "Идет авторизация. "
                                 "Просьба не закрывать браузер, "
                                 "даже если авторизация уже завершилась/не потребовалась. "
                                 "Проверка может занять некоторое время...")
        authorization = threading.Thread(target=self.authorization_thread, args=(), daemon=True)
        self.widget_disabled()
        authorization.start()

    def authorization_thread(self):
        result = self.whatsapp.authorization()
        logging.info(f"self.whatsapp.authorization() result {result}")
        self.authorization_end.emit(result)

    def set_authorization_label(self, check_result):
        logging.info("set_authorization_label")

        if not check_result:
            self.authorization_label.setText('<p style="color: red;">Требуется вход</p>')
            self.send_button.setEnabled(False)
        else:
            self.authorization_label.setText('<p style="color: green;">Вход выполнен</p>')
            self.send_button.setEnabled(True)

    def on_file_button_clicked(self):
        logging.info("on_file_button_clicked")

        default_path =  str(self.settings.value("excel/default_path"))
        file_path, _filter = QFileDialog.getOpenFileName(self, "Выбор файла",
                                                        default_path,
                                                         "*.xlsx")
        logging.info(f"chose file: {file_path} default_path = {default_path}")
        self.file_line.setText(file_path)

    def read_excel(self):
        logging.info("read_excel")

        read_excel = threading.Thread(target=self.excel_reader.read_excel, args=(self.file_line.text(),), daemon=True)
        self.widget_disabled()
        read_excel.start()

    @Slot(array)
    def read_excel_end(self, phones):
        self.phones = phones
        self.phone_count_label.setText(str(len(self.phones)))
        self.widget_enabled()

    def on_test_button_clicked(self):
        logging.info("on_test_button_clicked")
        self.send("test")

    def on_send_button_clicked(self):
        logging.info("on_send_button_clicked")
        self.send("whatsapp")

    def send(self, mode):
        logging.info(f"send mode = {mode}")
        if len(self.phones) == 0:
            logging.info("phones is empty")
            QMessageBox.critical(self, "Не загружены телефоны",
                                 "Некому отправить такую прекрасную рассылку. "
                                 "Выберете какой-нибудь файл excel, где есть колонка с именем Телефон "
                                 "и в этой колонке какие-нибудь телефоны")
            return
        elif self.spam_text.toPlainText() == "":
            logging.info("text is empty")
            QMessageBox.critical(self, "Пустой текст рассылки",
                                 "Давайте не будем такое отправлять. Напишите что-нибудь")
            return

        answer = QMessageBox.question(
            self,
            'Подумайте еще разок',
            'Вы точно хотите это отправить ' + self.phone_count_label.text() + ' контактам?',
            QMessageBox.StandardButton.Yes |
            QMessageBox.StandardButton.No
        )
        if answer == QMessageBox.StandardButton.Yes:
            logging.info("start spam")
            self.logging(mode, "Рассылка начата... Пути назад нет :)")
            self.widget_disabled()
            spam = threading.Thread(target=self.send_thread, args=(mode,), daemon=True)
            spam.start()
        else:
            logging.info("no spam")
            self.logging(mode, "На нет и суда нет :)")

    def send_thread(self, mode):
        logging.info(f"send_thread mode = {mode}")
        for phone in self.phones:
            log = None
            logging.debug(f"send mode = {mode} phone = {phone}")
            self.log_signal.emit(mode, "Попытка отправки сообщения на номер " + phone)
            if mode == "test":
                log = self.test_spam(phone, self.spam_text.toPlainText())
            elif mode == "whatsapp":
                log = self.whatsapp.send_spam(phone, self.spam_text.toPlainText())

            logging.debug(f"sent mode = {mode} phone = {phone}")
            self.log_signal.emit(mode, log)
            sleep(self.send_delay.value())

        logging.info(f"send_thread end mode = {mode}")
        self.spam_sent.emit(mode)

    def test_spam(self, phone, text):
        logging.info(f"test_spam phone = {phone}, text = {text}")
        return "[" + str(phone) + "]: " + str(text)

    @Slot(bool)
    def authorization_whatsapp_end(self, result):
        self.set_authorization_label(result)
        self.widget_enabled()


    @Slot(str)
    def send_whatsapp_end(self, mode):
        logging.info(f"widget_enabled mode = {mode}")
        if mode == "whatsapp":
            logging.info("widget_enabled close browser")
            self.whatsapp.close_browser()

        self.widget_enabled()
        self.logging(mode, "Выдыхаем, все закончилось")

    def widget_enabled(self):
        self.test_button.setEnabled(True)
        self.file_button.setEnabled(True)
        self.send_button.setEnabled(True)
        self.authorization_button.setEnabled(True)
        self.spam_text.setEnabled(True)
        self.load_gif.movie().stop()
        self.load_gif.setVisible(False)

    def widget_disabled(self):
        self.test_button.setDisabled(True)
        self.file_button.setDisabled(True)
        self.send_button.setDisabled(True)
        self.authorization_button.setDisabled(True)
        self.spam_text.setDisabled(True)
        self.load_gif.movie().start()
        self.load_gif.setVisible(True)

    @Slot(str, str)
    def logging(self, mode, text):
        logging.info(f"logging mode = {mode}, text = {text}")
        self.log.appendHtml('<p>' + color_mode(mode) + "[" + str(datetime.now()) + "] " + str(text) + '</p')
        self.log.repaint()



app = QApplication([])
window = WhatsappSpamWindow()
window.show()
sys.exit(app.exec())
