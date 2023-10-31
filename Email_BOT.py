# %%
import sys
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QTextEdit, QPushButton
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# %%
class EmailSender(QThread):
    update_signal = pyqtSignal(int)

    def __init__(self, from_email, password_email, to_email, subject, message, parent=None):
        super().__init__(parent)
        self.running = False
        self.from_email = from_email
        self.password_email = password_email
        self.to_email = to_email
        self.subject = subject
        self.message = message

    def run(self):
        self.running = True
        self.successful_emails = 0

        smtp_server = "smtp.office365.com"
        smtp_port = 587
        smtp_username = self.from_email
        smtp_password = self.password_email

        while self.running:
            msg = MIMEMultipart()
            msg["From"] = smtp_username
            msg["To"] = self.to_email
            msg["Subject"] = self.subject
            msg.attach(MIMEText(self.message, "plain"))

            try:
                with smtplib.SMTP(smtp_server, smtp_port) as server:
                    server.starttls()
                    server.login(smtp_username, smtp_password)
                    server.sendmail(smtp_username, self.to_email, msg.as_string())

                print("Email sent successfully!")
                self.successful_emails += 1
                self.update_signal.emit(self.successful_emails)

            except Exception as e:
                print("An error occurred:", str(e))

        print("Sending stopped.")

class EmailBot(QMainWindow):

    def __init__(self):
        super().__init__()

        self.init_ui()
        self.sender_thread = None

    def init_ui(self):
        self.setWindowTitle("Email Bot Oficce365")
        self.setGeometry(100, 100, 400, 500)
        self.setStyleSheet("background-color: #282729;")

        self.label_description = QLabel("Email Bot foi criado com o intuito de dispararmos e-mails para os\ncordenadores do SENAC, afim de solucionar nossos problemas.\nO desenvolvedor não se responsabiliza pelas ações do usúario!\n\nUtilizar somente Hotmail ou Outlook!", self)
        self.label_description.setGeometry(20, 20, 350, 75)
        self.label_description.setStyleSheet("color: #F04;")

        self.label_from = QLabel("From:", self)
        self.label_from.move(20, 120)
        self.label_from.setStyleSheet("color: #D9D9D9;")

        self.from_edit = QLineEdit(self)
        self.from_edit.setGeometry(100, 120, 280, 25)

        self.label_password = QLabel("Password:", self)
        self.label_password.move(20, 160)
        self.label_password.setStyleSheet("color: #D9D9D9;")

        self.password_edit = QLineEdit(self)
        self.password_edit.setGeometry(100, 160, 280, 25)

        self.label_to = QLabel("To:", self)
        self.label_to.move(20, 200)
        self.label_to.setStyleSheet("color: #D9D9D9;")

        self.to_edit = QLineEdit(self)
        self.to_edit.setGeometry(100, 200, 280, 25)

        self.label_subject = QLabel("Subject:", self)
        self.label_subject.move(20, 240)
        self.label_subject.setStyleSheet("color: #D9D9D9;")

        self.subject_edit = QLineEdit(self)
        self.subject_edit.setGeometry(100, 240, 280, 25)

        self.label_message = QLabel("Message:", self)
        self.label_message.move(20, 280)
        self.label_message.setStyleSheet("color: #D9D9D9;")

        self.message_edit = QTextEdit(self)
        self.message_edit.setGeometry(20, 310, 360, 100)

        self.send_button = QPushButton("Start Sending", self)
        self.send_button.setGeometry(140, 420, 120, 30)
        self.send_button.clicked.connect(self.toggle_sending)
        self.send_button.setStyleSheet("color: #D9D9D9;")

        self.success_label = QLabel("Successful Emails: 0", self)
        self.success_label.setGeometry(20, 460, 180, 20)
        self.success_label.setStyleSheet("color: #4F0;")

        self.sending = False

    def toggle_sending(self):
        if not self.sending:
            self.sending = True
            self.send_button.setText("Stop Sending")
            self.start_sending()

        else:
            self.sending = False
            self.send_button.setText("Start Sending")
            self.stop_sending()

    def start_sending(self):
        if not self.sender_thread:
            from_email = self.from_edit.text()
            password_email = self.password_edit.text()
            to_email = self.to_edit.text()
            subject = self.subject_edit.text()
            message = self.message_edit.toPlainText()

            self.sender_thread = EmailSender(from_email, password_email, to_email, subject, message)
            self.sender_thread.update_signal.connect(self.update_successful_emails)
            self.sender_thread.start()

    def stop_sending(self):
        if self.sender_thread:
            self.sender_thread.running = False
            self.sender_thread = None

    def update_successful_emails(self, count):
        self.success_label.setText(f"Successful Emails: {count}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EmailBot()
    window.show()
    sys.exit(app.exec())



