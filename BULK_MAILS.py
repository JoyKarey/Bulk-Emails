import pandas as pd
import os
import win32com.client as win32

def send_emails_with_attachments_from_excel(excel_path, pdf_dir, subject_template, signature_image_path, cc_addresses=None):
    # Load the Excel file containing institution names and emails
    df = pd.read_excel(excel_path)

    # Start Outlook application
    outlook = win32.Dispatch('Outlook.Application')

    # Initialize report lists
    sent_emails = []
    failed_emails = []

    for _, row in df.iterrows():
        institution = row['Institution Name']
        email = row['Email']

        # Create a new email
        mail = outlook.CreateItem(0)  # 0 indicates a mail item
        mail.Subject = subject_template.format(institution=institution)

        # Attach the signature image
        attachment = mail.Attachments.Add(signature_image_path)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "image001.jpg")

        # Construct the HTML body with formatting and signature
        mail.HTMLBody = (
            f"<html><body>"
            f"<p style='font-family:Calibri; font-size:11pt; color:black;'>Greetings,</p><br>"
            f"<p style='font-family:Calibri; font-size:11pt; color:black;'>Hope this finds you well.</p><br>"
            f"<p style='font-family:Calibri; font-size:11pt; color:black;'>#BODY.</p><br>"
            f"<p style='font-family:Calibri; font-size:11pt; color:black;'>#BODY.</p><br>"
            f"<p style='font-family:Calibri; font-size:11pt; color:black;'>#BODY</p><br>"
            f"<p style='font-family:Times New Roman; font-size:12pt; color:#7030A0; font-weight:bold;'>Regards,</p>"
            f"<p style='font-family:Times New Roman; font-size:12pt; color:#7030A0; font-weight:bold;'>#NAME</p>"
            f"<p style='font-family:Times New Roman; font-size:12pt; color:#7030A0; font-weight:bold;'>#DESIGNATION</p>"
            f"<p style='font-family:Times New Roman; font-size:12pt; color:#7030A0; font-weight:bold;'>#YOUR ADDRESS</p>"
            f"<p style='font-family:Times New Roman; font-size:12pt; color:#7030A0; font-weight:bold;'>Office: +254 #YOUR NUMBER|</p>"
            f"<p style='font-family:Times New Roman; font-size:12pt; color:#7030A0; font-weight:bold;'>Email: <a href='#YOUR EMAIL'>#INSTITUTION EMAIL</a></p>"
            f"<img src='cid:image001.jpg' alt='Signature'><br>"
            f"</body></html>"
        )
        mail.To = email

        # Add CC recipients if any
        if cc_addresses:
            mail.CC = ";".join(cc_addresses)

        # Attach the PDF for the current institution
        pdf_file = os.path.join(pdf_dir, f"{institution}.pdf")
        if os.path.exists(pdf_file):
            mail.Attachments.Add(pdf_file)
            mail.Send()  # Uncomment to actually send the email
            sent_emails.append((institution, email, "Sent"))
        else:
            failed_emails.append((institution, email, "Failed - No attachment"))

    # Create a report DataFrame
    report_df = pd.DataFrame(sent_emails + failed_emails, columns=['Institution Name', 'Email', 'Status'])
    report_path = os.path.join(pdf_dir, 'Email_Report.xlsx')
    report_df.to_excel(report_path, index=False)
    print(f"Email report saved to {report_path}")

# Define email templates
subject_template = "SUBJECT"

# Paths to the Excel file with institution names and emails, PDF directory, and signature image
excel_path = r'C:\Users\JWANGARI\Desktop\PYTHON\Test.xlsx'#PATH TO YOUR EXCEL FILE CONTAINING THE INSTITUTION NAMES AND EMAILS
pdf_dir = r'C:\Users\JWANGARI\Desktop\PYTHON\SCANS'#PATH CONTAINING PDFS
signature_image_path = r'C:\Users\JWANGARI\Desktop\PYTHON\sign.png'#PATH CONTAINING AN IMAGE SIGNITAURE
cc_addresses = ['Joy Wangari <joykamau32@gmail.com>']# ANY ADDRESSES YOU WOULD LIKE TO CC

# Execute the function
send_emails_with_attachments_from_excel(excel_path, pdf_dir, subject_template, signature_image_path, cc_addresses)
