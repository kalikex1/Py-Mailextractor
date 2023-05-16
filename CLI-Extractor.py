import os
import re
import pandas as pd

def extract_emails_from_csv(csv_file):
    emails = set()
    with open(csv_file, 'r') as file:
        reader = pd.read_csv(file)
        for column in reader.columns:
            for item in reader[column]:
                # Extract emails using regular expression
                email_matches = re.findall(r'[\w\.-]+@[\w\.-]+', str(item))
                emails.update(email_matches)
    return emails

def extract_emails_from_excel(excel_file):
    emails = set()
    reader = pd.read_excel(excel_file, engine='openpyxl')
    for column in reader.columns:
        for item in reader[column]:
            # Extract emails using regular expression
            email_matches = re.findall(r'[\w\.-]+@[\w\.-]+', str(item))
            emails.update(email_matches)
    return emails

def process_csv_folder(folder_path):
    output_file = os.path.join(folder_path, 'email_addresses.txt')

    emails = set()
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if filename.endswith('.csv'):
            extracted_emails = extract_emails_from_csv(file_path)
        elif filename.endswith('.xlsx') or filename.endswith('.xls'):
            extracted_emails = extract_emails_from_excel(file_path)
        else:
            continue
        emails.update(extracted_emails)

    with open(output_file, 'w') as file:
        file.write('\n'.join(emails))

    print(f"Emails extracted and saved to '{output_file}'.")

# Set the folder path
folder_path = input("Enter the path to the folder containing the CSV and Excel files: ")

# Process the CSV folder
process_csv_folder(folder_path)
