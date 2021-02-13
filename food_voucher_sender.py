from email.mime.text import MIMEText
from datetime import timedelta
import sys
import smtplib
import time
import xlrd

def open_data(source):
    # Open Excel document
    
    wb = xlrd.open_workbook(source)
    data_sheet = wb.sheet_by_index(0)

    return data_sheet

def get_values(data_sheet, email_col, a_col, b_col):
    # Extract email and codes or email, user, and password, add to dictionary of email:[code] email:[user,pass] pairs. 
    # e.g. {a@gmail.com:['CODE'], b@yahoo.co.uk:['CODE','CODE']}
    
    send_total = data_sheet.nrows-1
    count = 1  # Start at 1 as first row is titles
    stored = ''
    email_vals = {}

    for i in range(send_total):
        eAddr = data_sheet.cell_value(count, email_col)
        if eAddr == '':  # If blank, use previous Email
            eAddr = stored
        if b_col:
            username = data_sheet.cell_value(count,a_col)
            password = data_sheet.cell_value(count, b_col)
            if eAddr in email_vals: 
                email_vals[eAddr].append(username)  # If email already exists, append new user pass pair so one email for multiple users and pass.
                email_vals[eAddr].append(password)  # Two appends instead of tuple as allows code reuse later.
            else:
                email_vals[eAddr] = [username, password]
        else:
            code = data_sheet.cell_value(count, a_col)  # cell_value(ROW,COL)
            if eAddr in email_vals: 
                email_vals[eAddr].append(code)  # If email already exists, append new code so one email for multiple codes
            else:
                email_vals[eAddr] = [code]          
        stored = eAddr  # Stored email incase next blank therefore same family
        count += 1

    return email_vals

def format_email(email_vals, sender, pwd, subject, multi):
    # Format email using MIME. 
    # Function includes 1 time based and 1 exponential rate throttler.
    print('-- Debug -- in format')
    count = 1
    throttle = 1
    missed = []
    start = time.time()
    limit_timer = start
    
    for addr in email_vals:
        if count % 31 == 0:
            loop_time = time.time() - limit_timer
            if int(loop_time) < 60:
                delay = int(60-loop_time)
                print(f'!!! Hit rate limit. Sleeping for {delay} sec(s) !!!')
                time.sleep(delay)
                print('Continuing...')
                limit_timer = time.time()
        details = ''
        email = addr
        e_values = email_vals.get(addr, '')
        for i in e_values:
            if multi:
                details = f'{details+str(i)}\n'
            else:
                details = f'{details+str(i)}\n\n'
            

        body = f'EMAIL BODY HERE\n\nCodes: {details}'
        msg = MIMEText(body)
        msg['To'] = email
        msg['From'] = sender
        msg['Subject'] = subject
        
        try:
            send_email(sender, pwd, msg)
            print(f"Email sent to: {email}")  # Printing to terminal decreases send speed.
            print(details)
            print("*"*20)
        except KeyboardInterrupt:
            sys.exit()
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
    
    return print(f'Sent {len(email_vals)} emails in {time_taken}')
    
def send_email(sender, pwd, msg):
    # Send MIME formatted email
    print('-- Debug -- in send')
    svr = 'smtp.office365.com'
    port = '587'
    server = smtplib.SMTP(svr, port)
    
    server.ehlo()
    server.starttls()
    server.login(sender, pwd)
    server.send_message(msg)
    server.quit()
    
    return

def food_voucher_sender(source, sender, pwd, subject):
    email_col = 3  # D
    code_col = 7   # H
    unused_col = None  # Unused in this function
    
    data_sheet = open_data(source)
    print('-- Debug -- got sheet')
    if (data_sheet.cell_value(0, email_col) != 'Email' or 
            data_sheet.cell_value(0, code_col) != 'Code'):
        return print('ERROR: Column name values do not match.')
    email_codes = get_values(data_sheet, email_col, code_col, unused_col)
    print('-- Debug -- got vals')
    format_email(email_codes, sender, pwd, subject, False)
    return
    
def user_pass_sender(source, sender, pwd, subject):
    email_col = 3
    user_col = 6  # G
    pwd_col = 7   # H
    data_sheet = open_data(source)
    if (data_sheet.cell_value(0, email_col) != 'Email' or 
            data_sheet.cell_value(0, user_col) != 'Username'  or 
            data_sheet.cell_value(0, pwd_col) != 'Password'):
        return print('ERROR: Column name values do not match.')
    e_usr_pass = get_values(data_sheet, email_col, user_col, pwd_col)
    format_email(e_usr_pass, sender, pwd, subject, True)
    return
    
def main():
    # Office 365 imposes a limit of
    # 30 messages sent per minute, and a limit of 10,000 recipients per day.
    
    source = 'code_sample.xlsx'
    sender = ''  # Outlook Email address here. e.g. test@outlook.com
    pwd = ''  # Plaintext password here
    subject = 'User pass'  # Constant between every email
    
    selection = 'fv'
    
    if selection == 'fv':
        food_voucher_sender(source, sender, pwd, subject)
    elif selection == 'ep':
        user_pass_sender(source, sender, pwd, subject)
    else:
        return print('Error')
    return
    
if __name__ == "__main__":
    main()