while True:
    directory_path = select_file(title="Please select a file from the directory.")

    directory = directory_path.split('/')
    del directory[-1]
    directory = '/'.join(directory)
    directory = directory + '/'
    filenames = os.listdir(directory)
    
    if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
        break
    else:
        show_message('Error', 'Please re-enter the directory.')

filenames


df = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Descriptions')
df.head()

Unit_number_list = list(df['Unit number'])
#print(Unit_number_list)

Description_list = list(df['Description'])
#print(Description_list)

product_descriptions = {}

for u, unit in enumerate(Unit_number_list):
	description = Description_list[u]
	product_descriptions[unit] = description

product_descriptions

df = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Prices')
df.head()

Unit_number_list = list(df['Unit number'])
#print(Unit_number_list)

price_list = list(df['Price per unit'])
#print(price_list)

product_prices = {}

for u, unit in enumerate(Unit_number_list):
	price = price_list[u]
	product_prices[unit] = price

product_prices


def get_order_data(invoice_num, filename, directory):

    workbook = openpyxl.load_workbook(directory+filename)
    filename = workbook['OrderForm']

    df1 = get_pdf(directory+'MasterInvoice.pdf')
    df1

    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)

    date = filename['C9'].value
    date = str(date).split('-')
    year = date[0]
    month = date[1]
    day = date[2][:2]
    date = f'{month}/{day}/{year}'
    #print(date)
    order_data['Date'] = date

    order_data['CustomerAddress1'] = filename['C16'].value
    order_data['CustomerName'] = filename['C14'].value
    order_data['CustomerCompany'] = filename['C15'].value

    city = filename['C17'].value
    state = filename['C18'].value
    zipcode = filename['C19'].value
    company_address2 = f'{city}, {state} {zipcode}'
    order_data['CustomerAddress2'] = company_address2

    price0 = product_prices.get(filename['B24'].value)
    price0str = f'${price0:,.2f}'
    order_data['Price0'] = price0str
    order_data['Quantity0'] = filename['C24'].value
    amount0 = product_prices.get(filename['B24'].value) * filename['C24'].value
    amount0str = f'${amount0:,.2f}'
    order_data['Amount0'] = amount0str

    price1 = product_prices.get(filename['B25'].value)
    quantity1 = filename['C25'].value
    if price1 is not None and quantity1 is not None:
        price1str = f'${price1:,.2f}'
        order_data['Price1'] = price1str
        order_data['Quantity1'] = quantity1

        amount1 = price1 * quantity1
        amount1str = f'${amount1:,.2f}'
        order_data['Amount1'] = amount1str

    price2 = product_prices.get(filename['B26'].value)
    quantity2 = filename['C26'].value
    if price2 is not None and quantity2 is not None:
        price2str = f'${price2:,.2f}'
        order_data['Price2'] = price2str
        order_data['Quantity2'] = quantity2

        amount2 = price2 * quantity2
        amount2str = f'${amount2:,.2f}'
        order_data['Amount2'] = amount2str
    
    order_data['Description0'] = product_descriptions.get(filename['B24'].value)
    orderdescription1 = product_descriptions.get(filename['B25'].value)
    if orderdescription1 is not None:
        order_data['Description1'] = product_descriptions.get(filename['B25'].value)
    orderdescription2 = product_descriptions.get(filename['B26'].value)
    if orderdescription2 is not None:
        order_data['Description2'] = product_descriptions.get(filename['B26'].value)

    if 'Amount0' in order_data:
        amount0str = int(order_data['Amount0'].replace('$', '').replace(',', '').replace('.', ''))
    if 'Amount1' in order_data:
        amount1str = int(order_data['Amount1'].replace('$', '').replace(',', '').replace('.', ''))
    if 'Amount2' in order_data:
        amount2str = int(order_data['Amount2'].replace('$', '').replace(',', '').replace('.', ''))

    if 'Amount0' in order_data and 'Amount1' in order_data and 'Amount2' in order_data:
        allamounts = amount0str + amount1str + amount2str
        allamountstr = f'${allamounts:,.2f}'
        order_data['Total'] = allamountstr
    elif 'Amount0' in order_data and 'Amount1' in order_data:
        twoamounts = amount0str + amount1str
        twoamountstr = f'${twoamounts:,.2f}'
        order_data['Total'] = twoamountstr
    elif 'Amount0' in order_data:
        oneamount = amount0str
        oneamountstr = f'${oneamount:,.2f}'
        order_data['Total'] = oneamountstr
    
    order_data
    
    customer_name = filename['C14'].value
    customer_email = filename['C13'].value
    invoice_date = date

    return order_data, customer_name, customer_email, invoice_date


def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf', directory+f'Invoice {invoice_num}.pdf', order_data)

gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):

    gmail_user = "vote4ray2005@gmail.com"

    # List of recipients

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Date variable

    date = order_data.get('Date')

    # Body of the email

    name = order_data.get('CustomerName')

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = email
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory+f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


filenames = os.listdir(directory)
filenames

invoice_num = 8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        invoice_num += 1
        test = get_order_data(invoice_num, filename, directory)
        order_data = test[0]
        name = test[1]
        email = test[2]
        date = test[3]
        #print(invoice_num)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)