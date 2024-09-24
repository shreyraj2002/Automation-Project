import imaplib
import email
from email.header import decode_header
import re
import openpyxl 
from datetime import datetime



#Before running the script changes to be made in this are - 
#1. Change the email id and password in the script
#2. Change the excel file path in the script
#3. Change the date mentioned in the script
#4. Change the subject keyword 




# Function to search for emails with a specific subject and date
def search_emails_with_subject_and_date(mail, subject_keyword, search_date):
    # Format the date to match email format
    formatted_date = search_date.strftime("%d-%b-%Y")
    # Search for emails on a specific date and with a particular subject
    result, data = mail.search(None, f'(ON "{formatted_date}" SUBJECT "{subject_keyword}")')
    if result == "OK":
        return data[0].split()  # Return list of email IDs
    else:
        return []

# Function to extract email addresses from email body
def extract_email_addresses(body):
    email_pattern = r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+'
    return re.findall(email_pattern, body)

# Function to fetch and process emails
def process_emails(mail, email_ids):
    emails = []
    for e_id in email_ids:
        result, msg_data = mail.fetch(e_id, "(RFC822)")
        if result == "OK":
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            # Decode the email subject
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding if encoding else "utf-8")

            # If the email message is multipart, extract its payload
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode("utf-8")
                        emails.extend(extract_email_addresses(body))
            else:
                body = msg.get_payload(decode=True).decode("utf-8")
                emails.extend(extract_email_addresses(body))

    return emails

# Function to write email addresses to Excel
def write_to_excel(email_list, filename="extracted_emails_gulati_mahima.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Emails"
    
    # Add headers
    ws.append(["Email Address"])

    # Write emails to the sheet
    for email_address in email_list:
        ws.append([email_address])
    
    wb.save(filename)

# Main script
def main():
    # Mailbox login credentials and IMAP server details
    username = "gulati.mahima@globalmasstransit.net"
    password = "1207#mahi@"
    imap_server = "mail.emailsrvr.com"  # e.g., 'imap.gmail.com' or 'imap.outlook.com'

    # Connect to the email server
    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(username, password)

    # Select the mailbox folder (usually "INBOX")
    mail.select("inbox")

    # Search for emails with specific subject and date
    subject_keyword = "Undelivered"
    search_date = datetime(2024, 9, 23)  # Example date from image

    # Fetch the list of email IDs matching the criteria
    email_ids = search_emails_with_subject_and_date(mail, subject_keyword, search_date)
    
    if email_ids:
        # Process the emails and extract email addresses
        extracted_emails = process_emails(mail, email_ids)
        
        # Write extracted emails to Excel
        write_to_excel(extracted_emails)
        print(f"Extracted {len(extracted_emails)} emails and saved to Excel.")
    else:
        print("No emails found with the given subject and date.")
    
    # Logout from the mail server
    mail.logout()

if __name__ == "__main__":
    main()
