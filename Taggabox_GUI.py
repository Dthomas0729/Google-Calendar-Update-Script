import tkinter
import sqlite3
import json
from openpyxl import Workbook, load_workbook
from datetime import datetime
import pandas as pd
from calendar_update import get_order, get_customer, Customer, RentalOrder
from woocommerce import API

wcapi = API(
    url="https://taggaboxmoving.com",
    consumer_key="ck_0a900a48cc070cb1b59ea65a5baca52ed1420a2a",
    consumer_secret="cs_b1cea84a2d95c7d2ebfcdff62e97bed06546d93d",
    version="wc/v3"
)

# COLLECT ORDER DATA
orders = wcapi.get('orders')
data = json.loads(json.dumps(orders.json(), indent=3))
print(data)

book = load_workbook(filename='Delivery order form-1.xlsx')
sheet1 = book.active
print(sheet1['B1'].value)

sheet1['G3'] = datetime.now()
book.save('Delivery order form-1.xlsx')


root = tkinter.Tk()
root.title('Taggabox File Management ')


def new_order():
    tkinter.Label(root, text='first name').grid(row=3)
    tkinter.Label(root, text='last name').grid(row=4)

    entry1 = tkinter.Entry(root)
    entry2 = tkinter.Entry(root)

    entry1.grid(row=3, column=1)
    entry2.grid(row=4, column=1)

    tkinter.Button(root, text='submit form!', command=submit).grid(row=5, column=1, pady=5)


def submit():
    pass


# Labels
menu_title = tkinter.Label(root, text='Taggabox File Management', font=('courier', 25, 'bold'))
menu_title.grid(row=0, column=0, columnspan=5, padx=200, pady=65)

# THIS CREATES THE BUTTONS FOR THE MAIN MENU
orders_button = tkinter.Button(root, text='RENTAL ORDERS', padx=20, pady=20, command=new_order)
employees_button = tkinter.Button(root, text='EMPLOYEES', padx=20, pady=20)
inventory_button = tkinter.Button(root, text='INVENTORY', padx=20, pady=20)

orders_button.grid(row=2, column=1)
employees_button.grid(row=2, column=2)
inventory_button.grid(row=2, column=3)


#root.mainloop()

