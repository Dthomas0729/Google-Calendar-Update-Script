from __future__ import print_function
from woocommerce import API
import json
import datetime
from datetime import datetime, timedelta
import pickle
import sqlite3
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


class Customer:
    def __init__(self, first=None, last=None, phone=None, email=None, delivery_address=None,
                 pickup_address=None, amount_paid=0):
        self.first = first
        self.last = last
        self.phone = phone
        self.email = email
        self.delivery_address = delivery_address
        self.pickup_address = pickup_address
        self.amount_paid = amount_paid

    @property
    def fullname(self):
        return f'{self.first} {self.last}'

    @fullname.setter
    def fullname(self, name):
        first, last = name.split(' ')
        self.first = first
        self.last = last

    def __repr__(self):
        return f'Customer({self.first}, {self.last}, {self.phone}, {self.email},' \
               f' {self.delivery_address}, {self.pickup_address}, {self.amount_paid})'

    def __str__(self):
        return f'''    {self.fullname}
phone: {self.phone}
email: {self.email}
delivery: {self.delivery_address}
pickup: {self.pickup_address}
'''

    def save(self):
        # CONNECT TO THE DATABASE AND CREATE A SERVER
        conn = sqlite3.connect('customers.db')
        c = conn.cursor()

        c.execute(f"INSERT INTO customers VALUES ('{self.first}','{self.last}','{self.phone}','{self.email}')")
        conn.commit()
        conn.close()


class RentalOrder(Customer):
    total_orders = 0

    def __init__(self, order_num, customer=None, lg_boxes=0, xl_boxes=0, lg_dollies=0,
                 xl_dollies=0, labels=0, zip_ties=0, bins=0, rental_period=0,
                 delivery_date=0, delivery_datetime=0, pickup_datetime=0):
        self.order_num = order_num
        self.customer = customer
        self.lg_boxes = int(lg_boxes)
        self.xl_boxes = int(xl_boxes)
        self.lg_dollies = int(lg_dollies)
        self.xl_dollies = int(xl_dollies)
        self.labels = int(labels)
        self.zip_ties = int(zip_ties)
        self.bins = int(bins)
        self.customer_signed = False
        self.employee_signed = False
        self.rental_period = rental_period
        self.delivery_date = delivery_date
        self.delivery_datetime = delivery_datetime
        self.pickup_datetime = pickup_datetime

    def __repr__(self):
        return f'RentalOrder({self.order_num}, {self.customer}, {self.lg_boxes}, {self.xl_boxes}, {self.lg_dollies},' \
               f'{self.xl_dollies}, {self.labels}, {self.zip_ties},' \
               f' {self.bins}, {self.customer_signed}. {self.employee_signed}, {self.rental_period},' \
               f' {self.delivery_datetime}, {self.pickup_datetime})'

    def __str__(self):
        return f''' 
Lg Boxes: {self.lg_boxes}
Xl Boxes: {self.xl_boxes}
Lg Dollies: {self.lg_dollies}
Xl Dollies: {self.xl_dollies}
Labels: {self.labels}
Zip-Ties: {self.zip_ties}
Bins: {self.bins}
Customer Signed: {self.customer_signed}
Employee Signed: {self.employee_signed}
Rental Period: {self.rental_period}
Delivery Date: {self.delivery_date}
Pick Up Date: {self.pickup_datetime}
'''

    def save(self):
        # CONNECT TO THE DATABASE AND CREATE A SERVER
        conn = sqlite3.connect('rental_orders.db')
        c = conn.cursor()

        c.execute(f'''INSERT INTO rental_orders VALUES 
('{self.order_num}','{self.customer}',
'{self.lg_boxes}','{self.xl_boxes}',
'{self.lg_dollies}', '{self.xl_dollies}',
'{self.labels}', '{self.zip_ties}',
'{self.bins}', '{self.customer_signed}',
'{self.employee_signed}','{self.rental_period}',
'{self.delivery_date}', '{self.pickup_datetime}')
''')
        conn.commit()
        conn.close()


def get_customer():
    global current_customer

    current_order = last_order
    f_name = current_order['billing']['first_name']
    l_name = current_order['billing']['last_name']
    phone = current_order['billing']['phone']
    email = current_order['billing']['email']
    address_1 = current_order['shipping']['address_1']
    address_2 = current_order['shipping']['address_2']
    city = current_order['shipping']['city']
    state = current_order['shipping']['state']
    postcode = current_order['shipping']['postcode']
    delivery_address = f'{address_1} {address_2}, {city}, {state} {postcode}'
    pickup_address = current_order['meta_data'][4]['value']

    current_customer = Customer(first=f_name, last=l_name, phone=phone, email=email,
                                delivery_address=delivery_address, pickup_address=pickup_address)

    # current_customer.save()
    print(json.dumps(current_order, indent=4))
    print(current_customer)


def get_order():
    global c_order
    current_order = last_order
    order_num = current_order['id']

    delivery_datetime = current_order['meta_data'][0]['value'] + current_order['meta_data'][1]['value']
    try:
        delivery_date = datetime.strptime(current_order['meta_data'][0]['value'], '%Y-%m-%d')
    except ValueError:
        delivery_date = datetime.strptime(current_order['meta_data'][0]['value'], '%m/%d/%Y')

    try:
        delivery_datetime = datetime.strptime(delivery_datetime, '%Y-%m-%d%H:%M')
    except ValueError:
        delivery_datetime = datetime.strptime(delivery_datetime, '%m/%d/%Y%H:%M %p')

    delivery_date = delivery_date.date()

    # THIS LOCATES THE RENTAL PERIOD AND CREATES PICKUP DATETIME OBJECT
    try:
        rental_period = int(current_order['line_items'][0]['meta_data'][0]['value'][0])
        if rental_period == 1:
            pickup_datetime = delivery_datetime + timedelta(days=7)
            rental_period = '1 Week'
        elif rental_period == 2:
            pickup_datetime = delivery_datetime + timedelta(days=14)
            rental_period = '2 Weeks'
        elif rental_period == 3:
            pickup_datetime = delivery_datetime + timedelta(days=21)
            rental_period = '3 Weeks'
        elif rental_period == 4:
            pickup_datetime = delivery_datetime + timedelta(days=28)
            rental_period = '4 Weeks'
    except IndexError:
        pickup_datetime = delivery_datetime + timedelta(days=7)
        rental_period = '1 Week'

    # THIS LOCATES WOO-COMMERCE PRODUCT ID AND CREATES BOX PACKAGE ACCORDINGLY
    # SETS VALUES FOR LG BOXES, XL BOXES, ETC
    for x in range(len(current_order['line_items'])):
        if current_order['line_items'][x]['product_id'] == 1270:
            lg_boxes = 70
            xl_boxes = 10
            lg_dollies = 4
            xl_dollies = 2
            labels = 80
            zip_ties = 80
            bins = 0
            break
        elif current_order['line_items'][x]['product_id'] == 1515:
            lg_boxes = 50
            xl_boxes = 10
            lg_dollies = 3
            xl_dollies = 1
            labels = 60
            zip_ties = 60
            bins = 0
            break
        elif current_order['line_items'][x]['product_id'] == 1510:
            lg_boxes = 35
            xl_boxes = 5
            lg_dollies = 2
            xl_dollies = 0
            labels = 40
            zip_ties = 40
            bins = 0
        elif current_order['line_items'][x]['product_id'] == 1505:
            lg_boxes = 18
            xl_boxes = 2
            lg_dollies = 1
            xl_dollies = 0
            labels = 20
            zip_ties = 20
            bins = 0
            break
        elif current_order['line_items'][x]['product_id'] == 1545:
            lg_boxes = 1
            xl_boxes = 0
            lg_dollies = 0
            xl_dollies = 0
            labels = 0
            zip_ties = 0
            bins = 0
            break
        elif current_order['line_items'][x]['product_id'] == 1291:
            lg_boxes = 0
            xl_boxes = 0
            lg_dollies = 0
            xl_dollies = 0
            labels = 0
            zip_ties = 0
            bins = current_order['line_items'][x]['quantity']
            break
        else:
            lg_boxes = 0
            xl_boxes = 0
            lg_dollies = 0
            xl_dollies = 0
            labels = 0
            zip_ties = 0
            bins = 0

    delivery_datetime = delivery_datetime.strftime('%Y-%m-%dT%H:%M:%S-04:00')
    pickup_datetime = pickup_datetime.strftime('%Y-%m-%dT%H:%M:%S-04:00')

    c_order = RentalOrder(order_num, lg_boxes=lg_boxes, xl_boxes=xl_boxes, lg_dollies=lg_dollies,
                          xl_dollies=xl_dollies, labels=labels, zip_ties=zip_ties, bins=bins,
                          rental_period=rental_period, delivery_date=delivery_date,
                          delivery_datetime=delivery_datetime, pickup_datetime=pickup_datetime)

    #c_order.save()

    print(c_order)


def post_events():

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    end_time = datetime.strptime(c_order.delivery_datetime, '%Y-%m-%dT%H:%M:%S-04:00') + timedelta(hours=1)
    end_time = end_time.strftime('%Y-%m-%dT%H:%M:%S-04:00')

    event = {
        'summary': f'{current_customer.fullname} Box Delivery',
        'location': f'{current_customer.delivery_address}',
        'description': f'''{current_customer.fullname}
phone: {current_customer.phone}
address: {current_customer.delivery_address}
email: {current_customer.email}
{c_order.lg_boxes} Lg Boxes
{c_order.xl_boxes} Xl Boxes
{c_order.lg_dollies} Dollies
{c_order.xl_dollies} Xl Dollies
{c_order.labels} Labels & Zip Ties
{c_order.bins} Bins''',

        'start': {
            'dateTime': f'{c_order.delivery_datetime}',
            'timeZone': 'America/New_York',
        },
        'end': {
            'dateTime': f'{end_time}',
            'timeZone': 'America/New_York',
        },
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 24 * 60},
                {'method': 'popup', 'minutes': 10},
            ],
        },
    }

    # Call the Calendar API
    event = service.events().insert(calendarId='primary', body=event).execute()

    end_time = datetime.strptime(c_order.pickup_datetime, '%Y-%m-%dT%H:%M:%S-04:00') + timedelta(hours=1)
    end_time = end_time.strftime('%Y-%m-%dT%H:%M:%S-04:00')

    pickup_event = {
        'summary': f'{current_customer.fullname} Box Pick Up',
        'location': f'{current_customer.pickup_address}',
        'description': f'''{current_customer.fullname}
phone: {current_customer.phone}
Address {current_customer.pickup_address}
email: {current_customer.email}
{c_order.lg_boxes} Lg Boxes                                                                      
{c_order.xl_boxes} Xl Boxes
{c_order.lg_dollies} Dollies
{c_order.xl_dollies} Xl Dollies
{c_order.labels} Labels & Zip Ties
{c_order.bins} Bins''',

        'start': {
            'dateTime': f'{c_order.pickup_datetime}',
            'timeZone': 'America/New_York',
        },
        'end': {
            'dateTime': f'{end_time}',
            'timeZone': 'America/New_York',
        },
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 24 * 60},
                {'method': 'popup', 'minutes': 10},
            ],
        },
    }
    pickup_event = service.events().insert(calendarId='primary', body=pickup_event).execute()
    print('Delivery Event created: %s' % (event.get('htmlLink')))
    print('Pick-up Event created: %s' % (pickup_event.get('htmlLink')))


def save_order(obj, filename):
    with open(filename, 'wb') as output:  # Overwrites any existing file.
        pickle.dump(obj, output, pickle.HIGHEST_PROTOCOL)


SCOPES = ['https://www.googleapis.com/auth/calendar']

# WOO-COMMERCE API CREDENTIALS
wcapi = API(
    url=os.environ.get('WCAPI_URL'),
    consumer_key=os.environ.get('WCAPI_CONSUMER_KEY'),
    consumer_secret=os.environ.get('WCAPI_CONSUMER_SECRET'),
    version="wc/v3"
)

# COLLECT ORDER DATA
orders = wcapi.get('orders')
data = orders.json()
last_order = data[0]
new_order = True


def main():

    while new_order:

        get_customer()
        get_order()
        with open('latest_order.pkl', 'rb') as o:
            if last_order == pickle.load(o):
                print('no recent orders to update')
                break
            else:
                save_order(last_order, 'latest_order.pkl')
                # post_events()
                break


if __name__ == '__main__':
    main()
