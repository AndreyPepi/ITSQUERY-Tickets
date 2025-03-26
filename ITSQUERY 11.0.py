import sys
import os
import json
import requests
from datetime import timedelta
from PySide6.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QDateEdit, QLineEdit, QFileDialog, QProgressBar
from PySide6.QtCore import QDate, Qt, QThread, Signal, QTimer
from PySide6.QtGui import QIcon

ORIGIN_DICT = {
    1: "Via web pelo cliente",
    2: "Via web pelo agente",
    3: "Recebido via email",
    4: "Gatilho do sistema",
    5: "Chat (online)",
    7: "Email enviado pelo sistema",
    8: "Formulário de contato",
    9: "Via web API",
    10: "Agendamento automático",
    11: "JiraIssue",
    12: "RedmineIssue",
    13: "ReceivedCall",
    14: "MadeCall",
    15: "LostCall",
    16: "DropoutCall",
    17: "Acesso remoto",
    18: "WhatsApp",
    19: "MovideskIntegration",
    20: "ZenviaChat",
    21: "NotAnsweredCall",
    22: "FacebookMessenger",
    23: "WhatsApp Business Movidesk",
    24: "Altu",
    25: "WhatsApp Ativo"
}

TYPE_DICT = {
    1: "internal",
    2: "public"
}

def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class TokenInputDialog(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ITSQUERY Tickets")
        self.setWindowIcon(QIcon(get_resource_path("logo_mini.png")))
        self.resize(300, 100)

        self.token_label = QLabel("Digite o Token da API:")
        self.token_input = QLineEdit()
        self.confirm_button = QPushButton("Confirmar")
        self.confirm_button.clicked.connect(self.confirm_token)

        layout = QVBoxLayout()
        layout.addWidget(self.token_label)
        layout.addWidget(self.token_input)
        layout.addWidget(self.confirm_button)
        self.setLayout(layout)

    def confirm_token(self):
        token = self.token_input.text().strip()
        if token:
            self.close()
            self.main_window = HelpdeskQueryApp(token)
            self.main_window.show()

def ensure_list_length(lst, length, default="NA"):
    return lst + [default] * (length - len(lst)) if len(lst) < length else lst[:length]

class ConsultaThread(QThread):
    consulta_concluida = Signal(list)
    erro_consulta = Signal(str)

    def __init__(self, api_key, start_date, end_date):
        super().__init__()
        self.api_key = api_key
        self.start_date = start_date
        self.end_date = end_date

    def run(self):
        api_url = (f"https://api.movidesk.com/public/v1/tickets/past?token={self.api_key}"
                   f"&$select=id,subject,baseStatus,origin,createdDate,resolvedIn,chatWaitingTime,chatTalkTime,"
                   f"lifetimeWorkingTime,serviceFull,type,urgency,status"
                   f"&$expand=owner($select=businessName),clients($expand=organization($select=businessName)),"
                   f"customFieldValues($expand=items),actions($expand=timeAppointments($select=workTime))"
                   f"&$filter=createdDate ge {self.start_date}T00:00:00Z and createdDate le {self.end_date}T23:59:59Z")

        response = requests.get(api_url)

        if response.status_code == 200:
            data = response.json()

            filtered_data = [
                {
                    "urgency": ticket.get("urgency"),
                    "type": TYPE_DICT.get(ticket.get("type"), "Desconhecido"),
                    "Service1": ensure_list_length(ticket.get("serviceFull", []), 3)[0],
                    "Service2": ensure_list_length(ticket.get("serviceFull", []), 3)[1],
                    "Service3": ensure_list_length(ticket.get("serviceFull", []), 3)[2],
                    "lifeTimeWorkingTime": ticket.get("lifeTimeWorkingTime"),
                    "chatTalkTime": ticket.get("chatTalkTime"),
                    "chatWaitingTime": ticket.get("chatWaitingTime"),
                    "resolvedIn": ticket.get("resolvedIn"),
                    "owner": ticket.get("owner")["businessName"] if isinstance(ticket.get("owner"), dict) else "Desconhecido",
                    "organization": (
                        ticket.get("clients", [{}])[0].get("organization")["businessName"]
                        if isinstance(ticket.get("clients", [{}])[0].get("organization"), dict) and "businessName" in ticket.get("clients", [{}])[0].get("organization")
                        else "Desconhecido"
                    ),
                    "clients": ticket.get("clients", [{}])[0].get("businessName", "Desconhecido"),
                    "createdDate": ticket.get("createdDate"),
                    "origin": ORIGIN_DICT.get(ticket.get("origin"), "Desconhecido"),
                    "baseStatus": ticket.get("baseStatus"),
                    "subject": ticket.get("subject"),
                    "id": ticket.get("id"),
                }
                for ticket in data
            ]

            self.consulta_concluida.emit(filtered_data)
        else:
            self.erro_consulta.emit(f"Erro {response.status_code}: {response.text}")

class HelpdeskQueryApp(QWidget):
    def __init__(self, api_key):
        super().__init__()
        self.api_key = api_key
        self.initUI()

    def initUI(self):
        self.setWindowTitle("ITSQUERY Tickets")
        self.setWindowIcon(QIcon(get_resource_path("logo_mini.png")))

        self.start_date_label = QLabel("Data de Início:")
        self.start_date_input = QDateEdit()
        self.start_date_input.setCalendarPopup(True)
        self.start_date_input.setDate(QDate.currentDate())

        self.end_date_label = QLabel("Data Final:")
        self.end_date_input = QDateEdit()
        self.end_date_input.setCalendarPopup(True)
        self.end_date_input.setDate(QDate.currentDate())

        self.status_label = QLabel("Aguardando consulta...")
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setTextVisible(False)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_fake_progress)

        self.submit_button = QPushButton("Consultar")
        self.submit_button.clicked.connect(self.submit)

        self.save_button = QPushButton("Salvar Como")
        self.save_button.setVisible(False)
        self.save_button.clicked.connect(self.save_as)

        self.new_query_button = QPushButton("Nova Consulta")
        self.new_query_button.setVisible(False)
        self.new_query_button.clicked.connect(self.new_query)

        self.resize(300, 150)

        layout = QVBoxLayout()
        layout.addWidget(self.start_date_label)
        layout.addWidget(self.start_date_input)
        layout.addWidget(self.end_date_label)
        layout.addWidget(self.end_date_input)
        layout.addWidget(self.status_label)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.submit_button)
        layout.addWidget(self.save_button)
        layout.addWidget(self.new_query_button)

        self.setLayout(layout)

    def submit(self):
        self.status_label.setText("Consultando tickets...")
        self.progress_bar.setValue(0)
        self.timer.start(100)

        self.start_date = self.start_date_input.date().toString("yyyy-MM-dd")
        self.end_date = self.end_date_input.date().toString("yyyy-MM-dd")

        self.thread = ConsultaThread(self.api_key, self.start_date, self.end_date)
        self.thread.consulta_concluida.connect(self.on_consulta_concluida)
        self.thread.erro_consulta.connect(self.on_erro_consulta)
        self.thread.start()

        self.submit_button.setVisible(False)

    def update_fake_progress(self):
        if self.progress_bar.value() < 70:
            self.progress_bar.setValue(self.progress_bar.value() + 5)
        else:
            self.timer.stop()

    def on_consulta_concluida(self, data):
        self.timer.stop()
        self.status_label.setText("Consulta concluída!")
        self.progress_bar.setValue(100)
        self.consulta_data = data 
        self.save_button.setVisible(True)
        self.new_query_button.setVisible(True)

    def on_erro_consulta(self, erro_msg):
        self.timer.stop()
        self.status_label.setText(f"Erro: {erro_msg}")
        self.progress_bar.setValue(0)
        self.submit_button.setVisible(True)

    def save_as(self):
        start_date_str = self.start_date_input.date().toString("dd.MM.yyyy")
        end_date_str = self.end_date_input.date().toString("dd.MM.yyyy")

        default_filename = f"Relatório_{start_date_str}-{end_date_str}.json"

        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Salvar Como", default_filename, "JSON Files (*.json);;All Files (*)", options=options)
        
        if file_name:
            try:
                with open(file_name, 'w', encoding='utf-8') as file:
                    json.dump(self.consulta_data, file, ensure_ascii=False, indent=4)
                self.status_label.setText("Arquivo salvo com sucesso!")
            except Exception as e:
                self.status_label.setText(f"Erro ao salvar arquivo: {str(e)}")



    def new_query(self):
        self.close()
        self.token_window = TokenInputDialog()
        self.token_window.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    token_window = TokenInputDialog()
    token_window.show()
    sys.exit(app.exec())

# Versão 11.0
# Desenvolvido por Andrey Malinovski Pepi