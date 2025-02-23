import sys
import os
import json
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QDateEdit, QMessageBox, QFileDialog, QDialog, QCheckBox
)
from PySide6.QtCore import Qt, QDate
from mt5_init import connect_to_mt5
from get_deals import download_real_transactions, export_transactions_to_excel

class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Login Information")
        self.layout = QVBoxLayout(self)

        self.login_edit = QLineEdit()
        self.login_edit.setPlaceholderText("Login")
        self.layout.addWidget(QLabel("Login:"))
        self.layout.addWidget(self.login_edit)

        self.password_edit = QLineEdit()
        self.password_edit.setPlaceholderText("Password")
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.layout.addWidget(QLabel("Password:"))
        self.layout.addWidget(self.password_edit)

        self.server_edit = QLineEdit()
        self.server_edit.setPlaceholderText("Server")
        self.layout.addWidget(QLabel("Server:"))
        self.layout.addWidget(self.server_edit)

        self.save_button = QPushButton("Save")
        self.save_button.clicked.connect(self.save_info)
        self.layout.addWidget(self.save_button)

    def save_info(self):
        data = {
            "login": self.login_edit.text(),
            "password": self.password_edit.text(),
            "server": self.server_edit.text()
        }
        with open("log.json", "w") as f:
            json.dump(data, f)
        QMessageBox.information(self, "Info", "Login info saved.")
        self.accept()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MT5 Transactions")
        self.resize(400, 350)

        self.transactions = []

        layout = QVBoxLayout()
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        self.login_button = QPushButton("Login")
        self.login_button.clicked.connect(self.handle_login)
        layout.addWidget(self.login_button)

        date_layout = QHBoxLayout()
        self.start_date = QDateEdit(calendarPopup=True)
        self.start_date.setDate(QDate.currentDate())
        self.end_date = QDateEdit(calendarPopup=True)
        self.end_date.setDate(QDate.currentDate())

        date_layout.addWidget(QLabel("Start Date:"))
        date_layout.addWidget(self.start_date)
        date_layout.addWidget(QLabel("End Date:"))
        date_layout.addWidget(self.end_date)
        layout.addLayout(date_layout)

        self.include_opening_checkbox = QCheckBox("Include Opening Trades")
        self.include_opening_checkbox.setChecked(True)
        layout.addWidget(self.include_opening_checkbox)

        self.download_button = QPushButton("Download Transactions")
        self.download_button.clicked.connect(self.download_transactions)
        layout.addWidget(self.download_button)

        self.export_button = QPushButton("Export to Excel")
        self.export_button.clicked.connect(self.export_transactions)
        layout.addWidget(self.export_button)

    def handle_login(self):
        if os.path.exists("log.json"):
            with open("log.json", "r") as f:
                creds = json.load(f)
            if connect_to_mt5(creds["login"], creds["password"], creds["server"]):
                QMessageBox.information(self, "Success", "Connected to MT5")
            else:
                QMessageBox.warning(self, "Error", "Failed to connect to MT5")
        else:
            dialog = LoginDialog(self)
            dialog.exec()

    def download_transactions(self):
        if not os.path.exists("log.json"):
            QMessageBox.warning(self, "Error", "No login information found.")
            return

        with open("log.json", "r") as f:
            creds = json.load(f)

        if not connect_to_mt5(creds["login"], creds["password"], creds["server"]):
            QMessageBox.warning(self, "Error", "Failed to connect to MT5")
            return

        start = self.start_date.date().toPython()
        end = self.end_date.date().toPython()
        include_opening = self.include_opening_checkbox.isChecked()

        self.transactions = download_real_transactions(start, end, include_opening)
        QMessageBox.information(self, "Downloaded", f"{len(self.transactions)} transactions downloaded")

    def export_transactions(self):
        if not self.transactions:
            QMessageBox.warning(self, "Error", "No transactions to export.")
            return

        export_transactions_to_excel(self.transactions)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
