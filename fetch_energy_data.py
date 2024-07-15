import requests
import json
import pandas as pd
import csv


# response = requests.get("https://api.energy-charts.info/public_power?country=de")
# url = "https://api.coincap.io/v2/assets"
url = "https://api.energy-charts.info/public_power?country=de"
headers = {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
}
response = requests.request("GET", url, headers=headers, data={})
myJson = response.json()
# data = response.json()
# print(myJson)
arrdata = []
csvheader =['SYMBOL', 'NAME', 'PRICE(USD)']

unix = myJson['unix_seconds']
for x in myJson['production_types']:
    # listing = [x['symbol'], x['name'], x['priceUsd']]
    listing = [x['name'], x['data']]
    arrdata.append(listing)

    with open('auto3.csv', 'w', encoding='UTF8', newline='') as f:
        writer = csv.writer(f)

        writer.writerow(csvheader)
        writer.writerows(arrdata)
# print(type(data))

print('done')


# print(len(data['production_types']))
# print(len(data['unix_seconds']))

# index = data['unix_seconds']
# data_list = [
#     {'name': production_type['name'], 'data': production_type['data']}
#     for production_type in data['production_types']
# ]

# df = pd.DataFrame(data_list, index=index)  


# df = df.transpose()

# print(df)




# # df = pd.DataFrame(data)
# # df2 = pd.json_normalize(data)
# # print('df starts here', df)

# # Normalize the nested JSON structure into a flat table
# # flattened_data = []
# # for chart_data in data:  # Iterate through the list of chart data
# #     for entry in chart_data['data']:
# #         date = entry  # Extract the date from the entry
# #         for source in chart_data['source_data']:
# #             flattened_data.append({
# #                 "Date": date,
# #                 "Source": source['type'],
# #                 "Forecast": source['forecast'],
# #                 "Actual": source['actual']
# #             })

# # df = pd.DataFrame(flattened_data)

# # excel_file = 'energy_data.xlsx'
# # sheet_name = 'Sheet1'

# # try:
# #     # Try to load an existing workbook
# #     book = load_workbook(excel_file)
# #     writer = pd.ExcelWriter(excel_file, engine='openpyxl')
# #     writer.book = book
# # except FileNotFoundError:
# #     # Create a new workbook if it doesn't exist
# #     writer = pd.ExcelWriter(excel_file, engine='openpyxl')
# #     book = Workbook()
# #     writer.book = book

# # df.to_excel(writer, sheet_name=sheet_name, index=False)

# # writer.save()
# # writer.close()

# # print("Data has been written to", excel_file)
