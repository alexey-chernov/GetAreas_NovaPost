import requests
import openpyxl

# Ваш API ключ
API_KEY = '58ab545f60529346b896f526b30d1ad4'

# URL для запитів до API Нової Пошти
API_URL = "https://api.novaposhta.ua/v2.0/json/"

# Функція для отримання областей
def get_areas():
    payload = {
        "apiKey": API_KEY,
        "modelName": "Address",
        "calledMethod": "getAreas",
        "methodProperties": {}
    }
    response = requests.post(API_URL, json=payload).json()
    return response['data']

# Функція для отримання міст по області
def get_cities(area_ref):
    payload = {
        "apiKey": API_KEY,
        "modelName": "Address",
        "calledMethod": "getCities",
        "methodProperties": {
            "AreaRef": area_ref
        }
    }
    response = requests.post(API_URL, json=payload).json()
    return response['data']

# Функція для отримання відділень по місту
def get_warehouses(city_ref):
    payload = {
        "apiKey": API_KEY,
        "modelName": "AddressGeneral",
        "calledMethod": "getWarehouses",
        "methodProperties": {
            "CityRef": city_ref
        }
    }
    response = requests.post(API_URL, json=payload).json()
    return response['data']

# Створюємо Excel файл і записуємо дані
def save_to_excel(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Nova Poshta"

    # Заголовки
    ws.append(["Область", "Місто", "Відділення"])

    for area, cities in data.items():
        for city, warehouses in cities.items():
            for warehouse in warehouses:
                ws.append([area, city, warehouse])

    wb.save("nova_poshta_data.xlsx")

# Основна функція
def main():
    all_data = {}
    areas = get_areas()

    for area in areas:
        print(area)
        area_name = area['Description']
        area_ref = area['Ref']
        cities = get_cities(area_ref)

        all_data[area_name] = {}

        for city in cities:
            print(city)
            city_name = city['Description']
            city_ref = city['Ref']
            warehouses = get_warehouses(city_ref)

            all_data[area_name][city_name] = [wh['Description'] for wh in warehouses]

    save_to_excel(all_data)

if __name__ == "__main__":
    main()
