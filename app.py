import imaplib
import email
from email.header import decode_header
from bs4 import BeautifulSoup
import streamlit as st
import pandas as pd

def connect_to_gmail(username, password):
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(username, password)
    return mail

def fetch_emails(mail, subject_keyword, fabric_keyword, num_emails):
    mail.select("inbox")
    result, data = mail.search(None, f'(SUBJECT "{subject_keyword}")')
    email_ids = data[0].split()

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

    emails.sort(reverse=True, key=lambda x: x[0])  # Sort by timestamp
    latest_emails = [email_id for _, email_id in emails[:num_emails]]

    results = []
    for email_id in latest_emails:
        res, msg = mail.fetch(email_id, "(RFC822)")
        for response_part in msg:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")

                subject = subject.replace('\n', ' ')
                company_name = subject.split("---")[1].strip() if "---" in subject else "Unknown"

                email_body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        if content_type == "text/html":
                            email_body = part.get_payload(decode=True).decode()
                            break
                else:
                    email_body = msg.get_payload(decode=True).decode()

                soup = BeautifulSoup(email_body, "html.parser")
                vendor_info = {}

                if fabric_keyword in subject:
                    vendor_info = extract_vendor_info(soup)
                    vendor_info['Company Name'] = 'Fabric Division'
                else:
                    vendor_info = extract_vendor_info(soup)
                    vendor_info['Company Name'] = company_name

                results.append(vendor_info)

    return results

def extract_vendor_info(soup):
    vendor_info = {
        'Vendor Code': soup.find(string="Vendor Code").find_next("td").text.strip(),
        'Vendor Name': soup.find(string="Vendor Name").find_next("td").text.strip(),
        'Bank Name': soup.find(string="Bank Name").find_next("td").text.strip(),
        'Account Number': soup.find(string="Account Number").find_next("td").text.strip(),
        'Amount Rs.': soup.find(string="Amount Rs.").find_next("td").text.strip(),
        'UTR Number': soup.find(string="UTR Number").find_next("td").text.strip(),
        'UTR Date': soup.find(string="UTR Date").find_next("td").text.strip(),
        'Invoice Details': []
    }

    invoice_table = soup.find(string="Invoice Details:").find_next("table")
    for row in invoice_table.find_all("tr")[1:]:  # Skip the header row
        columns = row.find_all("td")
        if columns:
            invoice_info = {
                'Invoice No': columns[0].text.strip(),
                'Invoice Date': columns[1].text.strip(),
                'Paid Amount': columns[2].text.strip(),
                'TDS Amount': columns[3].text.strip(),
                'Supplier Invoice Amount': columns[4].text.strip()
            }
            vendor_info['Invoice Details'].append(invoice_info)

    return vendor_info

# Streamlit UI
st.title("Lakshmi Crane Service")

username = st.text_input("Gmail Username")
password = st.text_input("Gmail Password", type="password")
subject_keyword = "WE HAVE MADE RTGS FOR SUPPLIER DUE LIST ON"
fabric_keyword = "Fabric Division"
num_emails = st.number_input("Number of emails to fetch", min_value=1, max_value=100)

st.markdown("""
    <style>
    .stButton>button {
        background-color: green;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)

if st.button("Fetch Data"):
    if username and password:
        mail = connect_to_gmail(username, password)
        results = fetch_emails(mail, subject_keyword, fabric_keyword, num_emails)

        for vendor_info in results:
            st.subheader("Vendor Information")
            for key, value in vendor_info.items():
                if key == 'Invoice Details':
                    # Convert invoice details to a DataFrame and display as a table
                    invoice_df = pd.DataFrame(value)
                    st.write("Invoice Details:")
                    st.table(invoice_df)
                else:
                    st.write(f"{key}: {value}")

        mail.logout()
    else:
        st.error("Please enter your Gmail credentials.")

st.write("## Contact Details")
st.write("### Anthonysamy K ")
st.write("**Phone :** 9942024792")
st.write("**Address :** Andalpuram , Rajapalayam")
st.markdown('**Instagram:** [@Lakshmi Crane Service](https://www.instagram.com/lakshmi_crane_services/)')
