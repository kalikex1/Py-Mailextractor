import os
import csv
import re

def extract_emails(csv_file):
    emails = set()
    with open(csv_file, 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            for item in row:
                # Extract emails using regular expression
                email_matches = re.findall(r'[\w\.-]+@[\w\.-]+', item)
                emails.update(email_matches)
    return emails

def process_csv_folder(folder_path):
    output_file = os.path.join(folder_path, 'email_addresses.txt')

    emails = set()
    for filename in os.listdir(folder_path):
        if filename.endswith('.csv'):
            csv_file = os.path.join(folder_path, filename)
            extracted_emails = extract_emails(csv_file)
            emails.update(extracted_emails)

    with open(output_file, 'w') as file:
        file.write('\n'.join(emails))

    print(f"Emails extracted and saved to '{output_file}'.")

# Set the folder path
folder_path = input("Enter the path to the folder containing the CSV files: ")

# Process the CSV folder
process_csv_folder(folder_path)