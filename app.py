import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# INSERT ANSWER FOR QUESTION 1 HERE
while True:
    directory = single_input("Input the directory with ProductData.xlsx, MasterInvoice.pdf, and Order files are located.")
    x = os.listdir(directory)
    try:    
        if 'ProductData.xlsx' in x and 'MasterInvoice.pdf' in x and 'Order1.xlsx' in x and 'Order2.xlsx' in x and 'Order3.xlsx' in x and 'Order4.xlsx' in x:
            if directory[-1] == '\\':
                break
            else:
                directory = directory + '\\'
                break
    except:
        show_message("Error", "Try again.")

print(directory)
os.listdir(directory)

# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}

descriptions = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Descriptions')
prices = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Prices')

for i, row in descriptions.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

for i, row in prices.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

product_prices

# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):
    order_details = {}
    order_sheet = openpyxl.load_workbook(f"{directory}{filename}")['OrderForm']
    row = 24
    count = 0
    total = 0
    while order_sheet[f'B{row}'].value is not None:
       order_details[f'Price{count}'] = f'${order_sheet[f'C{row}'].value}'
       order_details[f"Quantity{count}"] = order_sheet[f'B{row}'].value
       order_details[f'Amount{count}'] = f'${order_sheet[f'C{row}'].value * product_prices[order_sheet[f'B{row}'].value]:,.2f}'
       order_details[f'Description{count}'] = product_descriptions[order_sheet[f'B{row}'].value]
       total += order_sheet[f'C{row}'].value * product_prices[order_sheet[f'B{row}'].value]
       row += 1
       count += 1
    order_details['Total'] = f'${total:,.2f}'
    order_details['CustomerCompany'] = order_sheet['C15'].value
    order_details['CustomerAddress1'] = order_sheet['C16'].value
    order_details['CustomerAddress2'] = order_sheet['C17'].value + " " + order_sheet['C18'].value + ', ' + str(order_sheet['C19'].value)
    order_details['InvoiceNum'] = invoice_num
    name = order_sheet['C14'].value
    email = order_sheet['C13'].value
    date = order_sheet['C9'].value.strftime('%m/%d/%Y')
    order_details['CustomerName'] = name
    order_details['Date'] = date
    order_data = order_details
    return [order_data, name, email, date] 

order_data = get_order_data(8515, 'Order1.xlsx', directory)[0]
name = get_order_data(8515, 'Order1.xlsx', directory)[1]
email = get_order_data(8515, 'Order1.xlsx', directory)[2]
date = get_order_data(8515, 'Order1.xlsx', directory)[3]

print([name, email, date, order_data])

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf(f'{directory}MasterInvoice.pdf', f'{directory}Invoice {invoice_num}.pdf', order_data)

# INSERT ANSWER FOR QUESTION 5 HERE
gmail_password_path = select_file('Please select the gmail password file')

def send_email(invoice_num, name, email, date, directory):
    # Enter your gmail user name and password
    gmail_user = "ianjackson0719@gmail.com"
    gmail_password = open(gmail_password_path).read()
    # List of recipients
    recipients = [email]
    # Email subect
    subject = f'Your order from PythonTech Co. - Invoice {invoice_num}'
    # Body of the email
    body = f"""Dear {name},
    
Attached is your invoice dated {date}.
    
Thank you for your business.
    
PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555 - 2372
Email: sales@pythontechco.com"""

    # Initialize the email message
    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)
    # Add attachment
    path_to_file = f'{directory}Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
    # Log in to the Gmail server
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)
    # Send the email
    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, name, email, date, directory)

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
invoice_num = 8515
for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        get_order_data(invoice_num , file, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num += 1