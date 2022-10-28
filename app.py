import requests
import openpyxl

file = "excel.xlsx"
from_row = 0
to_row = 0
origin_column = "A"
target_column = "B"
API_key = "YOUR_GOOGLE_MAPS_API_KEY"

book = openpyxl.load_workbook(file)
sheet = book.active


for i in range(from_row, to_row):
    try:
        origin = sheet[f'{origin_column}{i}'].value
        destination = sheet[f'{target_column}{i}'].value
        if not origin or not destination:
            break
        url = f"https://maps.googleapis.com/maps/api/distancematrix/json?origins={origin}&destinations={destination}&units=metric&key={API_key}"
        payload = {}
        headers = {}
        response = requests.request("GET", url, headers=headers, data=payload)
        sheet[f'C{i}'] = response.json()["rows"][0]["elements"][0]["distance"]["text"]
    except:
        pass


book.save(file)
