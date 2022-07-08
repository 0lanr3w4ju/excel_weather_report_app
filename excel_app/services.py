from flask import send_from_directory
import requests
from openpyxl import load_workbook


def weather(city):

    url = "https://community-open-weather-map.p.rapidapi.com/weather"
    querystring = {"q":f"{city}","lat":"0","lon":"0","callback":"test","id":"2172797","lang":"null","units":"imperial","mode":"HTML"}

    headers = {
        "X-RapidAPI-Key": "0a109dce92msh070cc32fb93c003p1e4e8djsnef8cf195fed1",
        "X-RapidAPI-Host": "community-open-weather-map.p.rapidapi.com"
    }

    response = requests.request("GET", url, headers=headers, params=querystring)

    return response.text

def process_xl(filename):

    wb = load_workbook(filename=filename)

    sheet_1 = wb['Sheet1']
    sheet_2 = wb['Sheet2']

    city_1 = sheet_1['A1'].value
    city_2 = sheet_2['A1'].value

    city_1 = str(city_1)
    city_2 = str(city_2)

    sheet_1['C1'] = weather(city_1)
    sheet_2['C1'] = weather(city_2)

    wb.save(filename=filename)
    return send_from_directory('static/files/', 'TestXl.xlsx')


    
