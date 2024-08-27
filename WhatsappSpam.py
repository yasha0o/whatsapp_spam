import sys
import threading
from datetime import datetime
from time import sleep

from PyQt6.QtWidgets import (
    QApplication,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QPlainTextEdit,
    QSplitter,
    QWidget, QLabel, QSpinBox, QLineEdit, QFileDialog, QMessageBox
)

import Whatsapp
from ReadExcel import read_excel
from PyQt6.QtCore import pyqtSignal as Signal, pyqtSlot as Slot, QSettings


class WatsappSpamWindow(QWidget):

    spam_sent = Signal(str)
    log_signal = Signal(str, str)
    settings = QSettings("settings.ini", QSettings.Format.IniFormat)
    whatsapp = Whatsapp.Whatsapp(settings)

    def __init__(self):
        super().__init__()
        self.phones = []
        self.setWindowTitle("Рассылка сообщений в watsapp")

        self.file_line = QLineEdit()
        self.file_line.setReadOnly(True)
        self.file_line.textChanged.connect(self.read_excel)
        self.file_button = QPushButton("Выбрать файл...")
        self.file_button.clicked.connect(self.on_file_button_clicked)

        layout = QVBoxLayout()

        file_layout = QHBoxLayout()
        file_layout.addWidget(self.file_line)
        file_layout.addWidget(self.file_button)

        layout.addLayout(file_layout)

        self.phone_count_label = QLabel()

        download_layout = QHBoxLayout()
        download_layout.addWidget(QLabel("Загружено телефонов для рассылки:"))
        download_layout.addWidget(self.phone_count_label)
        download_layout.addWidget(QSplitter())

        layout.addLayout(download_layout)

        self.spam_text = QPlainTextEdit()

        text_layout = QVBoxLayout()
        text_layout.addWidget(QLabel("Текст рассылки:"))
        text_layout.addWidget(self.spam_text)

        layout.addLayout(text_layout)

        self.send_delay = QSpinBox()
        self.send_delay.setValue(5)
        self.test_button = QPushButton("Тест")
        self.test_button.clicked.connect(self.on_test_button_clicked)
        self.send_button = QPushButton("Отправить")
        self.send_button.clicked.connect(self.on_send_button_clicked)

        self.spam_sent.connect(self.widget_enabled)

        send_layout = QHBoxLayout()
        send_layout.addWidget(QLabel("Задержка отправки:"))
        send_layout.addWidget(self.send_delay)
        send_layout.addWidget(QLabel("сек."))
        send_layout.addWidget(QSplitter())
        send_layout.addWidget(self.test_button)
        send_layout.addWidget(self.send_button)

        layout.addLayout(send_layout)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log_signal.connect(self.logging)

        log_layout = QVBoxLayout()
        log_layout.addWidget(QLabel("Логирование"))
        log_layout.addWidget(self.log)

        layout.addLayout(log_layout)

        self.setLayout(layout)

    def on_file_button_clicked(self):
        file_path, _filter = QFileDialog.getOpenFileName(self, "Выбор файла",
                                                         self.settings.value("excel/default_path"),
                                                         "*.xlsx")
        print("choosed: " + file_path)
        self.file_line.setText(file_path)

    def read_excel(self):
        print("read excel")
        self.phones = read_excel(self.file_line.text())
        self.phone_count_label.setText(str(len(self.phones)))

    def on_test_button_clicked(self):
        self.send("test")

    def on_send_button_clicked(self):
        self.send("whatsapp")

    def send(self, mode):

        if (len(self.phones) == 0):
            QMessageBox.critical(self, "Не загружены телефоны",
                                 "Некому отправить такую прекрасную рассылку. "
                                 "Выберете какой-нибудь файл excel, где есть колонка с именем Телефон "
                                 "и в этой колонке какие-нибудь телефоны")
            return
        elif self.spam_text.toPlainText() == "":
            QMessageBox.critical(self, "Пустой текст рассылки", "Давайте не будем такое отправлять. Напишите что-нибудь")
            return

        self.test_button.setDisabled(True)
        self.file_button.setDisabled(True)
        self.send_button.setDisabled(True)
        self.spam_text.setDisabled(True)

        answer = QMessageBox.question(
            self,
            'Подумай еще разок',
            'Ты точно хочешь это отправить ' + self.phone_count_label.text() + ' контактам?',
            QMessageBox.StandardButton.Yes |
            QMessageBox.StandardButton.No
        )
        if answer == QMessageBox.StandardButton.Yes:
            if (mode == "whatsapp"):
                self.logging(mode, "Для начала авторизуйтесь в whatsapp")
                if self.whatsapp.authorization() == True:
                    self.logging(mode, "Авторизация успешна")
                else:
                    self.logging(mode, "Что-то пошло не так :(")
                    self.widget_enabled(mode)
                    return

            self.logging(mode, "Рассылка начата... Пути назад нет :)")
            spam = threading.Thread(target=self.send_thread, args=(mode,), daemon=True)
            spam.start()
        else:
            self.logging(mode, "На нет и суда нет :)")

    def send_thread(self, mode):
        for phone in self.phones:
            log = None
            if mode == "test":
                log = self.test_spam(phone, self.spam_text.toPlainText())
            elif mode == "whatsapp":
                log = self.whatsapp.send_spam(phone, self.spam_text.toPlainText())
            self.log_signal.emit(mode, log)
            sleep(self.send_delay.value())

        self.spam_sent.emit(mode)

    def test_spam(self, phone, text):
        return "[" + str(phone) + "]: " + str(text)

    @Slot(str)
    def widget_enabled(self, mode):
        if mode == "whatsapp":
            self.whatsapp.close_browser()

        self.logging(mode, "Выдыхаем, все закончилось")

        self.test_button.setEnabled(True)
        self.file_button.setEnabled(True)
        self.send_button.setEnabled(True)
        self.spam_text.setEnabled(True)

    @Slot(str, str)
    def logging(self, mode, text):
        self.log.appendPlainText("[" + str(mode) + "]" + "[" + str(datetime.now()) + "]" + str(text))

app = QApplication([])
window = WatsappSpamWindow()
window.show()
sys.exit(app.exec())
