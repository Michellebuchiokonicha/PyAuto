from openpyxl import Workbook
import real
import requests

api_url = "https://opendata.elia.be/api/explore/v2.1/catalog/datasets/ods087/records?limit=20"

response = requests.get(api_url)
if response.status_code == 200:
    data = response.json()

wb = Workbook()
ws = wb.active
ws.title = "Energy Consumption Data"

ws.append(["Datetime", "Region", "Monitored Capacity"])

for result in data["results"]:
    ws.append([result["datetime"], result["region"], result["monitoredcapacity"]])

    wb.save("energy_data.xlsx")

    print("Excel file created: energy_data.xlsx")
else:
    print(f"Error fetching data:  {response.status_code}")
    