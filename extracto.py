import os
import re
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets

class EmailExtractor(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CSV Email Extractor")
        self.folder_path = ""

        layout = QtWidgets.QVBoxLayout()
        self.folder_label = QtWidgets.QLabel("Select Folder:")
        layout.addWidget(self.folder_label)

        self.folder_button = QtWidgets.QPushButton("Browse")
        self.folder_button.clicked.connect(self.select_folder)
        layout.addWidget(self.folder_button)

        self.extract_button = QtWidgets.QPushButton("Extract Emails")
        self.extract_button.clicked.connect(self.process_csv_folder)
        layout.addWidget(self.extract_button)

        self.setLayout(layout)

    def select_folder(self):
        folder_dialog = QtWidgets.QFileDialog()
        folder_path = folder_dialog.getExistingDirectory(self, "Select Folder")
        if folder_path:
            self.folder_path = folder_path
            self.folder_label.setText("Selected Folder: " + folder_path)

    def extract_emails_from_csv(self, csv_file):
        emails = set()
        with open(csv_file, 'r') as file:
            reader = pd.read_csv(file)
            for column in reader.columns:
                for item in reader[column]:
                    email_matches = re.findall(r'[\w\.-]+@[\w\.-]+', str(item))
                    emails.update(email_matches)
        return emails

    def extract_emails_from_excel(self, excel_file):
        emails = set()
        reader = pd.read_excel(excel_file, engine='openpyxl')
        for column in reader.columns:
            for item in reader[column]:
                email_matches = re.findall(r'[\w\.-]+@[\w\.-]+', str(item))
                emails.update(email_matches)
        return emails

    def process_csv_folder(self):
        if self.folder_path:
            output_file = os.path.join(self.folder_path, 'email_addresses.txt')

            emails = set()
            for filename in os.listdir(self.folder_path):
                file_path = os.path.join(self.folder_path, filename)
                if filename.endswith('.csv'):
                    extracted_emails = self.extract_emails_from_csv(file_path)
                elif filename.endswith('.xlsx') or filename.endswith('.xls'):
                    extracted_emails = self.extract_emails_from_excel(file_path)
                else:
                    continue
                emails.update(extracted_emails)

            with open(output_file, 'w') as file:
                file.write('\n'.join(emails))

            QtWidgets.QMessageBox.information(self, "Extraction Complete",
                                              f"Emails extracted and saved to '{output_file}'.")

# Create the application instance and run the event loop
app = QtWidgets.QApplication([])
window = EmailExtractor()
window.show()
app.exec_()