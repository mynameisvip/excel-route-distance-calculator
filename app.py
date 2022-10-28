import requests
import openpyxl

FILE = "excel.xlsx"
FROM_ROW = 0
TO_ROW = 0
ORIGIN_COLUMN = "A"
TARGET_COLUMN = "B"
API_KEY = "YOUR_GOOGLE_MAPS_API_KEY"

book = openpyxl.load_workbook(file)
sheet = book.active


for i in range(FROM_ROW, TO_ROW):
    try:
        origin = sheet[f'{ORIGIN_COLUMN}{i}'].value
        destination = sheet[f'{TARGET_COLUMN}{i}'].value
        if not origin or not destination:
            break
        url = f"https://maps.googleapis.com/maps/api/distancematrix/json?origins={origin}&destinations={destination}&units=metric&key={API_KEY}"
        response = requests.request("GET", url, headers={}, data={})
        sheet[f'C{i}'] = response.json()["rows"][0]["elements"][0]["distance"]["text"]
    except:
        pass


book.save(file)
