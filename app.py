from openpyxl import Workbook, load_workbook
import requests
import schedule
import time


import price_ahead
import wind_generation

schedule.every().hour.do(price_ahead.update_excel_data)
schedule.every().hour.do(wind_generation.update_excel_data)

while True:
    schedule.run_pending()
    time.sleep(1)