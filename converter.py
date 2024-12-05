import csv
import vobject
from datetime import datetime
import os
import re

def csv_to_vcards(csv_file_path, output_folder):
    with open(csv_file_path, mode='r', newline='', encoding='ISO-8859-1', errors='replace') as csv_file:
        reader = csv.DictReader(csv_file)
        for row in reader:
            try:
                print(f"Processing contact: {row.get('First Name', 'Unknown')} {row.get('Last Name', 'Unknown')}")

                vcard = vobject.vCard()

                # Name components
                vcard.add('n')
                vcard.n.value = vobject.vcard.Name(
                    family=row.get('Last Name', ''),
                    given=row.get('First Name', ''),
                    additional=row.get('Middle Name', ''),
                    prefix=row.get('Title', ''),
                    suffix=row.get('Suffix', '')
                )

                # Full name
                vcard.add('fn')
                vcard.fn.value = f"{row.get('Title', '')} {row.get('First Name', '')} {row.get('Middle Name', '')} {row.get('Last Name', '')} {row.get('Suffix', '')}".strip()

                # Organization and Job Title
                if row.get('Company') or row.get('Department'):
                    vcard.add('org')
                    vcard.org.value = [row.get('Company', ''), row.get('Department', '')]
                if row.get('Job Title'):
                    vcard.add('title')
                    vcard.title.value = row.get('Job Title')

                # Emails
                for i in range(1, 4):
                    email = row.get(f'E-mail {i} Address', '').strip()
                    if email:
                        email_field = vcard.add('email')
                        email_field.value = email
                        email_field.type_param = 'INTERNET'

                # Phone numbers
                phone_fields = {
                    'Primary Phone': 'VOICE',
                    'Home Phone': 'HOME',
                    'Home Phone 2': 'HOME',
                    'Mobile Phone': 'CELL',
                    'Pager': 'PAGER',
                    'Home Fax': 'FAX',
                    'Company Main Phone': 'WORK',
                    'Business Phone': 'WORK',
                    'Business Phone 2': 'WORK',
                    'Business Fax': 'FAX',
                    "Assistant's Phone": 'WORK',
                    'Other Phone': 'OTHER',
                    'Other Fax': 'FAX',
                    'Callback': 'CALLBACK',
                    'Car Phone': 'CAR',
                    'ISDN': 'ISDN',
                    'Radio Phone': 'RADIO',
                    'TTY/TDD Phone': 'TTY-TDD',
                    'Telex': 'TELEX'
                }
                for field, type_param in phone_fields.items():
                    phone_number = row.get(field, '').strip()
                    if phone_number:
                        tel = vcard.add('tel')
                        tel.value = phone_number
                        tel.type_param = type_param

                # Addresses
                address_types = {
                    'Home': ['Home Address', 'Home Street', 'Home Street 2', 'Home Street 3', 'Home Address PO Box', 'Home City', 'Home State', 'Home Postal Code', 'Home Country'],
                    'Work': ['Business Address', 'Business Street', 'Business Street 2', 'Business Street 3', 'Business Address PO Box', 'Business City', 'Business State', 'Business Postal Code', 'Business Country'],
                    'Other': ['Other Address', 'Other Street', 'Other Street 2', 'Other Street 3', 'Other Address PO Box', 'Other City', 'Other State', 'Other Postal Code', 'Other Country']
                }
                for addr_type, fields in address_types.items():
                    if any(row.get(field, '').strip() for field in fields):
                        address = vcard.add('adr')
                        address.type_param = addr_type.upper()
                        address.value = vobject.vcard.Address(
                            box=row.get(fields[4], '').strip(),
                            extended=row.get(fields[1], '').strip(),
                            street=row.get(fields[2], '').strip(),
                            city=row.get(fields[5], '').strip(),
                            region=row.get(fields[6], '').strip(),
                            code=row.get(fields[7], '').strip(),
                            country=row.get(fields[8], '').strip()
                        )

                # Save the vCard to a file
                contact_name = row.get('First Name', 'Unknown').replace(' ', '_')
                contact_name = re.sub(r'[^a-zA-Z0-9_]', '', contact_name)  # Remove invalid characters
                if not contact_name:
                    contact_name = "Unknown"
                output_file_path = f"{output_folder}/{contact_name}.vcf"
                with open(output_file_path, mode='w', encoding='utf-8') as vcard_file:
                    vcard_file.write(vcard.serialize())

            except Exception as e:
                print(f"Error processing contact: {row}")
                print(f"Exception: {e}")
                continue

if __name__ == "__main__":
    csv_file_path = 'YourPATH.csv'  # Replace with your CSV file path
    output_folder = 'vcards'  # Replace with your desired output folder
    os.makedirs(output_folder, exist_ok=True)  # Create the folder if it doesn't exist
    csv_to_vcards(csv_file_path, output_folder)
