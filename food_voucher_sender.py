from email.mime.text import MIMEText
from datetime import timedelta
import smtplib
import time
import xlrd

def open_data(source):
    # Open Excel document
    
    wb = xlrd.open_workbook(source)
    data_sheet = wb.sheet_by_index(0)

    return data_sheet

def get_values(data_sheet, email_col, code_col):
    # Extract email and codes, add to dictionary of email:[code] pairs. 
    # e.g. {a@gmail.com:['CODE'], b@yahoo.co.uk:['CODE','CODE']}
    
    amt_urls = data_sheet.nrows-1
    count = 1  # Start at 1 as first row is titles
    stored = ''
    emco_pairs = {}

    if data_sheet.cell_value(0,code_col) != 'Code' or data_sheet.cell_value(0,email_col) != 'Email':
        return("ERROR - Either Email or Code column mismatched.")
    for i in range(amt_urls):
        code = data_sheet.cell_value(count,code_col)  # cell_value(ROW,COL)
        eAddr = data_sheet.cell_value(count,email_col)
        if eAddr == '':  # If blank, use previous Email
            eAddr = stored
        if eAddr in emco_pairs: 
            emco_pairs[eAddr].append(code)  # If email already exists, append new code so one email for multiple codes
        else:
            emco_pairs[eAddr] = [code]          
        stored = eAddr  # Stored email incase next blank therefore same family
        count += 1

    return emco_pairs

def format_email(emco_pairs, sender, pwd, subject):
    # Format email using MIME. 
    # Function includes 1 time based and 1 exponential rate throttler.
    
    vc_tot = 0
    count = 1
    throttle = 1
    missed = []
    start = time.time()
    limit_timer = start
    
    for addr in emco_pairs:
        if count % 31 == 0:
            loop_time = time.time() - limit_timer
            if int(loop_time) < 60:
                delay = int(60-loop_time)/2
                print(f'!!! Hit rate limit. Sleeping for {delay} sec(s) !!!')
                time.sleep(delay)
                print('Continuing...')
                limit_timer = time.time()
        vc = ''
        email = addr
        code = emco_pairs.get(addr, '')
        for i in code:
            vc_tot += 1
            vc = f'{vc+str(i)}\n\n'

        body = f'EMAIL BODY HERE\n\nCodes: {vc}'
        msg = MIMEText(body)
        msg['To'] = email
        msg['From'] = sender
        msg['Subject'] = subject
        
        try:
            send_email(sender, pwd, msg)
            print(f"Email sent to: {email} with {len(code)} code(s)")  # Printing to terminal decreases send speed.
            print("*"*20)
        except:
            throttle += throttle
            print(f'!!! Hit exception. Sleeping for {throttle} secs. !!!')
            missed.append(msg)
            time.sleep(throttle)
            print("Continuing...")
        finally:
            count += 1

    print(f'{len(missed)} emails left due to raised exception.')
    if missed != None:
        for m in missed:
            try:
                send_email(sender, pwd, m)
            except:
                print(f"Unable to send email to: {m['To']}")
    time_taken = str(timedelta(seconds=round(time.time() - start, 2)))[2:10]
    
    return print(f'Sent {len(emco_pairs)} emails in {time_taken}')
    
def send_email(sender, pwd, msg):
    # Send MIME formatted email
    
    svr = 'smtp.office365.com'
    port = '587'
    server = smtplib.SMTP(svr, port)
    
    server.ehlo()
    server.starttls()
    server.login(sender, pwd)
    server.send_message(msg)
    server.quit()
    
    return
        

def main():
    # Office 365 imposes a limit of
    # 30 messages sent per minute, and a limit of 10,000 recipients per day.
    
    source = 'sample.xlsx'
    sender = ''  # Outlook Email address here. e.g. test@outlook.com
    pwd = ''  # Plaintext password here
    subject = 'Voucher Codes'  # Constant between every email
    email_col = 3  # D
    code_col = 7   # H
    
    data_sheet = open_data(source)
    email_codes = get_values(data_sheet, email_col, code_col)
    format_email(email_codes, sender, pwd, subject)
    

if __name__ == "__main__":
    main()