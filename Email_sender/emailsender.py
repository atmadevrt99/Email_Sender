import sys
import csv
from PyQt5.QtWidgets import *
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

class EmailSender(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Email Sender')
        self.setGeometry(100, 100, 750, 500)  # Adjusted size of the window

        self.tab_widget = QTabWidget(self)
        self.tab_widget.setGeometry(20, 20, 700, 450)  # Adjusted size of the tab widget

        self.tab_send = QWidget()
        self.tab_history = QWidget()

        self.tab_widget.addTab(self.tab_send, 'Send Email')
        self.tab_widget.addTab(self.tab_history, 'Email History')

        self.initSendTab()
        self.initHistoryTab()

    def initSendTab(self):
        self.lbl_subject = QLabel('Subject:', self.tab_send)
        self.lbl_subject.move(20, 20)
        self.txt_subject = QTextEdit(self.tab_send)
        self.txt_subject.setGeometry(20, 50, 450, 50)  # Adjusted width of the text area

        self.lbl_body = QLabel('Body:', self.tab_send)
        self.lbl_body.move(20, 110)
        self.txt_body = QTextEdit(self.tab_send)
        self.txt_body.setGeometry(20, 140, 450, 150)  # Adjusted width and height of the text area

        self.lbl_send_status = QLabel('', self.tab_send)  # Added label to display send status
        self.lbl_send_status.setGeometry(20, 310, 300, 30)  # Adjusted position and size of the label

        self.btn_upload = QPushButton('Upload CSV', self.tab_send)
        self.btn_upload.setGeometry(20, 350, 100, 30)
        self.btn_upload.clicked.connect(self.uploadCSV)

        self.btn_send = QPushButton('Send Email', self.tab_send)
        self.btn_send.setGeometry(150, 350, 100, 30)
        self.btn_send.clicked.connect(self.sendEmail)

        self.csv_file = ''

    def initHistoryTab(self):
        self.table_history = QTableWidget(self.tab_history)
        self.table_history.setGeometry(20, 20, 650, 320)  # Adjusted width and height of the table widget
        self.table_history.setColumnCount(5)  # Adjusted number of columns
        self.table_history.setHorizontalHeaderLabels(['Receiver Name', 'Email', 'Remarks', 'Sent Date', 'Sent Time'])  # Added column headers

        self.btn_generate_report = QPushButton('Generate Report', self.tab_history)
        self.btn_generate_report.setGeometry(20, 350, 120, 30)  # Adjusted position and size of the button
        self.btn_generate_report.clicked.connect(self.generateReport)

        self.lbl_report_status = QLabel('', self.tab_history)  # Added label to display report status
        self.lbl_report_status.setGeometry(150, 350, 300, 30)  # Adjusted position and size of the label

    def uploadCSV(self):
        options = QFileDialog.Options()
        self.csv_file, _ = QFileDialog.getOpenFileName(self, "Select CSV file", "", "CSV Files (*.csv)", options=options)

    def sendEmail(self):
        
        if not self.csv_file:
            self.lbl_send_status.setText('Please upload a CSV file.')  # Update send status label
            print('Please upload a CSV file.')
            return

        subject = self.txt_subject.toPlainText()
        body = self.txt_body.toPlainText()

        try:
            # SMTP Configuration
            smtp_server = 'smtp.gmail.com'
            smtp_port = 587
            smtp_username = 'your mail'
            smtp_password = 'your password'

            # Create SMTP connection
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(smtp_username, smtp_password)

            self.lbl_send_status.setText('Please wait... Sending emails')  # Indicate that emails are being sent

            # Read email addresses from CSV file and send emails
            with open(self.csv_file, 'r') as file:
                reader = csv.reader(file)
                next(reader)  # Skip header row
                for row_idx, row in enumerate(reader):
                    to_email = row[1]  # Assuming the second column contains email addresses

                    # Construct email message
                    msg = MIMEMultipart()
                    msg['To'] = to_email
                    msg['Subject'] = subject
                    msg.attach(MIMEText(body, 'plain'))

                    # Send email
                    server.sendmail(smtp_username, to_email, msg.as_string())

                    # Add entry to history tab with sent date and time
                    sent_date = datetime.now().strftime('%Y-%m-%d')
                    sent_time = datetime.now().strftime('%H:%M:%S')
                    self.table_history.insertRow(row_idx)
                    self.table_history.setItem(row_idx, 0, QTableWidgetItem(row[0]))  # Receiver Name
                    self.table_history.setItem(row_idx, 1, QTableWidgetItem(to_email))  # Email
                    self.table_history.setItem(row_idx, 2, QTableWidgetItem('Sent'))  # Remarks
                    self.table_history.setItem(row_idx, 3, QTableWidgetItem(sent_date))  # Sent Date
                    self.table_history.setItem(row_idx, 4, QTableWidgetItem(sent_time))  # Sent Time

            server.quit()

            self.lbl_send_status.setText('Emails sent successfully!')  # Update send status label
            # Clear text boxes after sending email
            self.txt_subject.clear()
            self.txt_body.clear()
        except Exception as e:
            self.lbl_send_status.setText(f'Failed to send emails: {e}')  # Update send status label


    def generateReport(self):
        if self.table_history.rowCount() == 0:
            self.lbl_report_status.setText('No data to generate report.')  # Update report status label
            return

        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Report", "", "CSV Files (*.csv)", options=options)

        if file_path:
            with open(file_path, 'w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(['Receiver Name', 'Email', 'Remarks', 'Sent Date', 'Sent Time'])
                for row_idx in range(self.table_history.rowCount()):
                    row_data = []
                    for col_idx in range(self.table_history.columnCount()):
                        item = self.table_history.item(row_idx, col_idx)
                        if item:
                            row_data.append(item.text())
                        else:
                            row_data.append('')
                    writer.writerow(row_data)

            self.lbl_report_status.setText('Report generated successfully.')  # Update report status label

if __name__ == '__main__':
    app = QApplication(sys.argv)
    sender = EmailSender()
    sender.show()
    sys.exit(app.exec_())
