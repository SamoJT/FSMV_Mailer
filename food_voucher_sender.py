from email.mime.text import MIMEText
from datetime import timedelta
import smtplib
import time
import xlrd
###################################################
# TO DO:
# Include rate limit check / catch to prevent code exiting due to hitting limit
###################################################

# Office 365 imposes a limit of
# 30 messages sent per minute, and a limit of 10,000 recipients per day.

def open_data():
    # Open Excel document with urls
    
    f = ('sample.xlsx')
    wb = xlrd.open_workbook(f)
    data_sheet = wb.sheet_by_index(0)
    return data_sheet

def get_values(data_sheet, email_col, code_col):
    # Extract email and codes, add to dictionary of email:[code] pairs. 
    # e.g. {a@gmail.com:['CODE'], b@yahoo.co.uk:['CODE','CODE']}
    
    amt_urls = data_sheet.nrows-1
    count = 1  # Start at 1 as first row is titles
    stored = ''
    email_urls = {}

    if data_sheet.cell_value(0,code_col) != 'Code' or data_sheet.cell_value(0,email_col) != 'Email':
        return("ERROR - Either Email or Code column mismatched.")
    for i in range(amt_urls):
        code = data_sheet.cell_value(count,code_col)  # cell_value(ROW,COL)
        eAddr = data_sheet.cell_value(count,email_col)
        if eAddr == '':  # If blank, use previous Email
            eAddr = stored
        if eAddr in email_urls: 
            email_urls[eAddr].append(code)  # If email already exists, append new code so one email for multiple codes
        else:
            email_urls[eAddr] = [code]          
        stored = eAddr  # Stored email incase next blank therefore same family
        count += 1
    return email_urls

def send_emails(email_urls, sender, pwd):
    # Create email and send to respective family.
    
    svr = 'smtp.office365.com'  # Connect constants
    port = '587'  # Connect constants
    subject = 'Voucher Codes'  # Constant between every email
    
    count = 0
    vc_tot = 0
    start = time.time()
    for addr in email_urls:
        vc = ''
        email = addr
        code = email_urls.get(addr, '')
        for i in code:
            vc_tot += 1
            vc = f'{vc+str(i)}\n\n'
        # Format email
        body = f'EMAIL BODY HERE\n\nCodes: {vc}'

        msg = MIMEText(body)
        msg['To'] = email
        msg['From'] = sender
        msg['Subject'] = subject

        # Send it #
        server = smtplib.SMTP(svr, port) #
        server.ehlo() #
        server.starttls() #
        server.login(sender, pwd) #
        server.send_message(msg) #
        print(f"Email sent to: {email} with {len(code)} code(s)")
        print("*"*20)
        server.quit() #
        count += 1
    end = str(timedelta(seconds=round(time.time() - start, 2)))[2:10]
    print(f'{vc_tot} codes sent to {count} emails in {end}')
    return
        

def main():
    sender = ''  # Outlook Email address here. e.g. test@outlook.com
    pwd = ''  # Plaintext password here
    email_col = 3  # D
    code_col = 7   # H
    
    data_sheet = open_data()
    values = get_values(data_sheet, email_col, code_col)
    email_urls = open_data()
    send_emails(email_urls, sender, pwd)

if __name__ == "__main__":
    main()