# COLLECT EMAIL LIST #

from woocommerce import API
import openpyxl
import os

#1. First I need to connect to woocommerce API
wcapi = API(
    url=os.environ.get('WCAPI_URL'),
    consumer_key=os.environ.get('WCAPI_CONSUMER_KEY'),
    consumer_secret=os.environ.get('WCAPI_CONSUMER_SECRET'),
    version="wc/v3"