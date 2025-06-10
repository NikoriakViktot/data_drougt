
import pandas as pd

import numpy as np
# Налаштуйте опції для відображення більшої кількості рядків
pd.set_option('display.max_rows', None)  # Відображення всіх рядків
pd.set_option('display.max_columns', None)  # Відображення всіх стовпців
pd.set_option('display.width', None)  # Автоматичне розширення ширини консолі
pd.set_option('display.max_colwidth', None)  # Показувати повний текст у стовпцях


file_error_stan = r"C:\Users\5302\PycharmProjects\data_drougt\ТСГ-6_2016_2016_Броди_06_2016_картопля.xls"
file_path = r'C:\Users\5302\PycharmProjects\data_drougt\DATA_base_soil_water_cleaned\ТСГ-6_2017\ТСГ-6_2017_Волинська\Любешів\Овес\ТСГ-6_2017_2017_Любешів_03_2017_овес.xls'
file_path1 = '.\ТСХ-6_Новодністровськ_07_2021_люцерна.xlsx'
file_path2 = '.\ТСХ-6_Чернівці_08_2021_соя.xlsx'
# file_date_nevalid=r"C:\Users\user\PycharmProjects\data_drought\DATA_base_soil_water\ТСГ-6_2016\ТСГ-6_2016_Івано-Франківська\Долина\зяб\ТСГ-6_Долина_08_2015_зяб.xls"
file_date_nevalid_dot = r"C:\Users\user\PycharmProjects\data_drought\DATA_base_soil_water\ТСГ-6_2020\ТСГ-6_2020_Харківська\Золочів\Горох\06_2020.xls"
def get_weather_station( data):
    # Визначаємо можливі позиції для інформації про станцію
    possible_positions = [3, 7]  # Індекси стовпців, де може знаходитись інформація
    station_name = ""

    for pos in possible_positions:
        content = data.iloc[4, pos]
        if pd.notna(content):  # Якщо вміст є
            content = content.strip()
            if 'Станція' in content:
                # Якщо слово 'Станція' в тому ж стовпці, що й назва
                station_name = content.replace('Станція', '').strip()
                break
            else:
                # Якщо 'Станція' не в цьому стовпці, використовуємо зміст як назву
                station_name = content

    return station_name if station_name else None

# Спробуємо відкрити файл
try:
    data = pd.read_excel(file_error_stan)


    print("\nВміст файлу:")
    # print(data)  # Виводить увесь вміст файла
    # print(data.iloc[0:78])
    print(data.iloc[0:78])
    print("\n NO VALID DATE MONTH")
    # print(data3.iloc[0:73, 0:2])
    print("\n NO VALID DATE DOT")
 # Виведе розміри датафрейму


    # print(station_name)




    # print(density_value)
except Exception as e:
    print(f"Помилка при відкритті файлу: {e}")

import xlrd


# workbook = xlrd.open_workbook(file_path)
# sheet = workbook.sheet_by_index(0)  # Вибираємо перший лист



workbook = xlrd.open_workbook(file_path)
sheets_data = []  # Лист для збереження даних з усіх аркушів

for sheet in workbook.sheets():
    sheet_data = []

    for row_idx in range(sheet.nrows):
        # Видалення пустих колонок у кожному рядку
        row_values = [cell for cell in sheet.row_values(row_idx) if cell != '']
        if row_values:  # Якщо рядок не порожній після очищення
            sheet_data.append(row_values)

    sheets_data.append(sheet_data)  # Додаємо очищені дані кожного аркуша до загального списку

# # Тепер sheets_data містить дані з усіх аркушів як список списків
# for sheet in sheets_data:
#     for row in sheet:
#         print(row)  # Виводить кожен рядок на новому рядку у терміналі
#     print("\n")  # Додає порожній рядок між листами для кращої читабельності
# text_10= """'Об\'ємна маса ґрунту на глибині ґрунту 10 см, г/ см³'"""
# text_20= """'Об\'ємна маса ґрунту на глибині ґрунту 20 см, г/ см³'"""
# text_30= """'Об\'ємна маса ґрунту на глибині ґрунту 30 см, г/ см³'"""
# text_40= """'Об\'ємна маса ґрунту на глибині ґрунту 40 см, г/ см³'"""
# text_50= """'Об\'ємна маса ґрунту на глибині ґрунту 50 см, г/ см³'"""
# text_60= """'Об\'ємна маса ґрунту на глибині ґрунту 60 см, г/ см³'"""
# text_70= """'Об\'ємна маса ґрунту на глибині ґрунту 70 см, г/ см³'"""
# text_80= """'Об\'ємна маса ґрунту на глибині ґрунту 80 см, г/ см³'"""
# text_90= """'Об\'ємна маса ґрунту на глибині ґрунту 90 см, г/ см³'"""
# text_100= """'Об\'ємна маса ґрунту на глибині ґрунту 100 см, г/ см³'"""
# # text_10= """'Об\'ємна маса ґрунту на глибині ґрунту 10 см, г/ см³'"""
# text_zap_neprod_10 = """Запаси непродуктивної вологи на глибині ґрунту 10 см, мм"""
# text_zap_prod_10 = """Запаси продуктивної вологи при НВ на глибині ґрунту 10 см, мм"""
# text_volo_10 = """Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 10 см, %"""
# text_zap_zag_10 = """Запаси загальної вологи по шарах на глибині ґрунту 10 см, мм"""
# text_zap_prod_shar_10 = """Запаси продуктивної вологи по шарах на глибині ґрунту 10 см, мм"""
# text_pidsum_10 = """Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 10 см, мм"""
# new_data = pd.DataFrame({
#     'Рік': [data.iloc[6, ]],
# 'Область': [data.iloc[6, ]],
# 'Метеостанція': [data.iloc[6, ]],
# 'Місяць': [data.iloc[6, ]],
# 'Дата': [data.iloc[6, ]],
# 'Культура': [data.iloc[6, ]],
# 'Попередня культура': [data.iloc[6, ]],
# 'Тип ґрунту': [data.iloc[6, ]],
# 'Ділянка': [data.iloc[6, ]],
# text_10: [data.iloc[6, ]],
# 'Рік': [data.iloc[6, ]],
# 'Рік': [data.iloc[6, ]],
# 'Рік': [data.iloc[6, ]],
# 'Рік': [data.iloc[6, ]],
# 'Рік': [data.iloc[6, ]],


