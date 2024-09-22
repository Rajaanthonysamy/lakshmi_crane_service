import re
import imaplib
import email
from email.header import decode_header
from bs4 import BeautifulSoup

def company_name_extractor(subject):
    subject = subject.replace('\n', ' ')
    print("Compa extrec com name is ",subject)
    # Use regex to extract the company name
    match = re.search(r'---(.+?)---', subject)
    if match:
        company_name = match.group(1).strip()
        print(f"Extracted Company Name: {company_name}")
    else:
        return "Not found check the bill"
    return company_name



# Connect to Gmail
username = "pass"
password = "pass"  # Consider using OAuth2 for better security
mail = imaplib.IMAP4_SSL("imap.gmail.com")

# Log in
mail.login(username, password)

# Select the mailbox you want to check
mail.select("inbox")

# Define your search criteria
subject_keyword = "WE HAVE MADE RTGS FOR SUPPLIER DUE LIST ON"
fabric_keyword = "Fabric Division"
num_emails = int(input("How many emails do you want to fetch? "))  # User input for number of emails

# Search for emails
result, data = mail.search(None, f'(SUBJECT "{subject_keyword}")')

# Get the list of email IDs
email_ids = data[0].split()

# Fetch the emails and store them with their date
emails = []
for email_id in email_ids:
    res, msg = mail.fetch(email_id, "(RFC822)")
    for response_part in msg:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            date_str = msg["Date"]
            date_tuple = email.utils.parsedate_tz(date_str)
            local_time = email.utils.mktime_tz(date_tuple)
            emails.append((local_time, email_id))

# Sort emails by date (newest first) and limit to the specified number
emails.sort(reverse=True, key=lambda x: x[0])  # Sort by timestamp
latest_emails = [email_id for _, email_id in emails[:num_emails]]

for email_id in latest_emails:
    # Fetch the email by ID
    res, msg = mail.fetch(email_id, "(RFC822)")
    for response_part in msg:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding if encoding else "utf-8")

    ##        print(f"\nSubject: {subject}")
            company_name= company_name_extractor(subject)

            # Print the email body (HTML)
            email_body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    if content_type == "text/html":
                        email_body = part.get_payload(decode=True).decode()
                        break  # Exit loop after finding HTML part
            else:
                email_body = msg.get_payload(decode=True).decode()

            # Print the email body
            # print("\nEmail Body:")
            # print(email_body)  # Print the raw HTML body

            # Use BeautifulSoup to parse the HTML content
            soup = BeautifulSoup(email_body, "html.parser")

            # Check if the email contains "Fabric Division"
            if fabric_keyword in subject:
                # Scrape vendor information for Fabric Division
                vendor_code = soup.find(string="Vendor Code").find_next("td").text.strip()
                vendor_name = soup.find(string="Vendor Name").find_next("td").text.strip()
                bank_name = soup.find(string="Bank Name").find_next("td").text.strip()
                account_number = soup.find(string="Account Number").find_next("td").text.strip()
                amount = soup.find(string="Amount Rs.").find_next("td").text.strip()
                utr_number = soup.find(string="UTR Number").find_next("td").text.strip()
                utr_date = soup.find(string="UTR Date").find_next("td").text.strip()

                # print("\nVendor Information:")
                # print(f"Vendor Code: {vendor_code}")
                # print(f"Vendor Name: {vendor_name}")
                # print(f"Bank Name: {bank_name}")
                # print(f"Account Number: {account_number}")
                # print(f"Amount Rs.: {amount}")
                # print(f"UTR Number: {utr_number}")
                # print(f"UTR Date: {utr_date}")

                # Extract invoice details
                invoice_table = soup.find(string="Invoice Details:").find_next("table")
                print("\n-------------------------------  Invoice Details -------------------------\n")
                for row in invoice_table.find_all("tr")[1:]:  # Skip the header row
                    columns = row.find_all("td")
                    if columns:
                        invoice_no = columns[0].text.strip()
                        invoice_date = columns[1].text.strip()
                        paid_amt = columns[2].text.strip()
                        tds_amt = columns[3].text.strip()
                        sup_inv_amt = columns[4].text.strip()
                        print(f"Company Name : Fabric Divison , Invoice No: {invoice_no}, Invoice Date: {invoice_date}, Paid Amount: {paid_amt}, TDS Amount: {tds_amt}, Supplier Invoice Amount: {sup_inv_amt}")

            else:
                # Scrape vendor information for other emails
                vendor_code = soup.find(string="Vendor Code").find_next().text.strip()
                vendor_name = soup.find(string="Vendor Name").find_next().text.strip()
                bank_name = soup.find(string="Bank Name").find_next().text.strip()
                account_number = soup.find(string="Account Number").find_next().text.strip()
                amount = soup.find(string="Amount Rs.").find_next().text.strip()
                utr_number = soup.find(string="UTR Number").find_next().text.strip()
                utr_date = soup.find(string="UTR Date").find_next().text.strip()

                # print("\nVendor Information:")
                # print(f"Vendor Code: {vendor_code}")
                # print(f"Vendor Name: {vendor_name}")
                # print(f"Bank Name: {bank_name}")
                # print(f"Account Number: {account_number}")
                # print(f"Amount Rs.: {amount}")
                # print(f"UTR Number: {utr_number}")
                # print(f"UTR Date: {utr_date}")

                # Extract invoice details
                invoice_table = soup.find(string="Invoice Details:").find_next("table")
                print("\n-------------------------------  Invoice Details -------------------------\n")
                for row in invoice_table.find_all("tr")[1:]:  # Skip the header row
                    columns = row.find_all("td")
                    if columns:
                        invoice_no = columns[0].text.strip()
                        invoice_date = columns[1].text.strip()
                        paid_amt = columns[2].text.strip()
                        tds_amt = columns[3].text.strip()
                        sup_inv_amt = columns[4].text.strip()
                        print(f"Company Name : {company_name} Invoice No: {invoice_no}, Invoice Date: {invoice_date}, Paid Amount: {paid_amt}, TDS Amount: {tds_amt}, Supplier Invoice Amount: {sup_inv_amt}")
                        print("------------------------------------------****--------------------------------------------")

# Log out
mail.logout()
