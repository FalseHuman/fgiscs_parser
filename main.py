import os
import json
import math
import requests
import openpyxl
from typing import List

fied_excel = {
    "name_career": "Наименование карьера/месторождения",
    "location": "Местонахождение карьера/месторождения (наименование ближайшего населенного пункта муниципального образования, адрес)",
    "countrySubjectId": "Субъект РФ",
    "latitude": "Географические координаты карьера/месторождения (широта)",
    "longitude": "Географические координаты карьера/месторождения (долгота)",
    "htmlAddress": "Адрес сайта",
    "name_companies": "Полное наименование",
    "inn": "ИНН",
    "kpp": "КПП",
    "contacts": "Контактные данные",
    "fullCode": "Код КСР",
    "name_ksrs": "Наименование",
    "unitName": "Единица измерения",
    "laywer_info": "Юридические лица (их обособленные подразделения), осуществляющие складирование / реализацию продукции",
    "materials_info": "Виды материалов"
}

def data_to_str(data: List[str], type_data: str) -> str:
    """
    Формируем строки из словаря
    """
    string = ""
    for item in data:
        for key in item.keys():
            if key != "id":
                if key == type_data:
                    string += "{}\n".format(item[key])
    return string[:-1]


def json_to_excel(json_path: str):
    with open(json_path) as f:
        data = json.load(f)
    wb = openpyxl.Workbook()
    # Формируем excel
    main_page = wb.create_sheet(
            title='Основная информация', index=0)
    sheet_main_page = wb['Основная информация']
    sheet_main_page.append([fied_excel["countrySubjectId"], fied_excel["name_career"], fied_excel["location"],\
                            fied_excel["latitude"], fied_excel["longitude"], fied_excel["htmlAddress"],\
                            fied_excel["name_companies"], fied_excel["inn"], fied_excel["kpp"],
                            fied_excel["contacts"], fied_excel["fullCode"], fied_excel["name_ksrs"], fied_excel["unitName"]
                        ])
    for item in data:
        name_companies, inn, kpp =  data_to_str(item["companies"], "name"), data_to_str(item["companies"], "inn"), data_to_str(item["companies"], "kpp") 
        contacts, fullCode, name_ksrs =  data_to_str(item["companies"], "contacts"), data_to_str(item["ksrs"], "fullCode"), data_to_str(item["ksrs"], "name")
        unitName = data_to_str(item["ksrs"], "unitName")
        sheet_main_page.append([item["countrySubjectId"], item["name"], item["location"],    
                                item["latitude"], item["longitude"], item["htmlAddress"], 
                                name_companies, inn, kpp,
                                contacts, fullCode, name_ksrs,
                                unitName
                                ])

    excel_file = "fgiscs_data.xlsx"
    wb.save(excel_file)


def parser_data(url: str) -> dict:
    """
    Парсим все данные с сайта, сначала получаем общее кол-во элементов
    """
    req_total = requests.get(url=url, verify=False)
    take = 100 # кол-во на спагинированной странице
    # Получаем кол-во записей
    total = req_total.json().get("total", 0)
    end_page = math.ceil(total / take)
    subjects_data = requests.get(url="https://fgiscs.minstroyrf.ru/api/Quarry/GetCountrySubjects", verify=False, timeout=30)
    data_items = []

    for page in range(1, end_page + 1):
        data = requests.get(url=url, params={"take": take, "page": page},  verify=False, timeout=30)
        items = data.json().get("items", [])
        for item in items:
            item["countrySubjectId"] = [data["name"] for data in subjects_data.json() if data["id"] == item["countrySubjectId"]][0]
            data_items.append(item)

    return data_items


if __name__ == '__main__':
    url = "https://fgiscs.minstroyrf.ru/api/Quarry/List"
    file_json = "result.json"
    data = parser_data(url)
    with open(file_json, "w", encoding="utf-8") as fp:
        json.dump(data, fp, ensure_ascii=False)
    json_to_excel(file_json)
    os.remove(file_json)