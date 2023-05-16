import os
import csv
import re
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog, QTextEdit, QVBoxLayout


class EmailExtractionGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('CSV Email Extraction')
        self.file_path_label = QLabel('Select CSV Folder:')
        self.file_path_line_edit = QLineEdit()
        self.browse_button = QPushButton('Browse')
        self.browse_button.clicked.connect(self.browse_folder)
        self.extract_button = QPushButton('Extract Emails')
        self.extract_button.clicked.connect(self.extract_emails)
        self.output_text_edit = QTextEdit()
        self.output_text_edit.setReadOnly(True)

        layout = QVBoxLayout()
        layout.addWidget(self.file_path_label)
        layout.addWidget(self.file_path_line_edit)
        layout.addWidget(self.browse_button)
        layout.addWidget(self.extract_button)
        layout.addWidget(self.output_text_edit)
        self.setLayout(layout)

    def browse_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, 'Select CSV Folder')
        self.file_path_line_edit.setText(folder_path)

    def extract_emails(self):
        folder_path = self.file_path_line_edit.text()
        if folder_path:
            emails = self.process_csv_folder(folder_path)
            self.output_text_edit.setPlainText('\n'.join(emails))
        else:
            self.output_text_edit.setPlainText('Please select a CSV folder.')

    def extract_emails_from_csv(self, csv_file):
        emails = set()
        with open(csv_file, 'r') as file:
            reader = csv.reader(file)
            for row in reader:
                for item in row:
                    # Extract emails using regular expression
                    email_matches = re.findall(r'[\w\.-]+@[\w\.-]+', item)
                    emails.update(email_matches)
        return emails

    def process_csv_folder(self, folder_path):
        emails = set()
        for filename in os.listdir(folder_path):
            if filename.endswith('.csv'):
                csv_file = os.path.join(folder_path, filename)
                extracted_emails = self.extract_emails_from_csv(csv_file)
                emails.update(extracted_emails)
        return emails


if __name__ == '__main__':
    app = QApplication([])
    gui = EmailExtractionGUI()
    gui.show()
    app.exec_()
