import MetaTrader5 as mt5
from datetime import datetime
import sys
import os
import json
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QDateEdit, QMessageBox, QFileDialog,
    QTableWidget, QTableWidgetItem, QDialog, QComboBox
)
from PySide6.QtCore import Qt, QDate
import openpyxl

# ---------------- MT5 Initialization ---------------- #

def connect_to_mt5(login, password, server):
    try:
        login_int = int(login)
    except ValueError:
        print("Invalid login: must be an integer.")
        return False

    if not mt5.initialize(login=login_int, password=password, server=server):
        print(f"Failed to connect to MT5: {mt5.last_error()}")
        return False

    print("Connected to MT5 successfully.")
    return True

# ---------------- Download Transactions ---------------- #

def download_real_transactions(start_date, end_date):
    deals = mt5.history_deals_get(start_date, end_date)
    if deals is None:
        print("No deals found.", mt5.last_error())
        return []

    transactions = []
    for deal in deals:
        trans = {
            "ticket": str(deal.ticket),
            "date": datetime.fromtimestamp(deal.time).strftime("%Y-%m-%d %H:%M:%S"),
            "type": "Buy" if deal.type == 0 else "Sell",
            "volume": deal.volume,
            "price": deal.price,
            "comment": deal.comment
        }
        transactions.append(trans)
    return transactions

# ---------------- Login Dialog ---------------- #

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

# ---------------- Main Window ---------------- #

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MT5 Transactions")
        self.resize(400, 300)

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

        self.download_button = QPushButton("Download Transactions")
        self.download_button.clicked.connect(self.download_transactions)
        layout.addWidget(self.download_button)

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
        if not mt5.initialize():
            QMessageBox.warning(self, "Error", "MT5 not connected")
            return

        start = self.start_date.date().toPython()
        end = self.end_date.date().toPython()

        self.transactions = download_real_transactions(start, end)
        QMessageBox.information(self, "Downloaded", f"{len(self.transactions)} transactions downloaded")

        mt5.shutdown()

# ---------------- Main Execution ---------------- #

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
