import logging
from pandas import to_datetime


def replace_date_format(date_text):
    if isinstance(date_text, datetime):
        return date_text  # If it's already a datetime object, return as is

    date_text = str(date_text).strip()

        # Including Ukrainian month names for parsing
    months_uk = {
        'січня': '01', 'лютого': '02', 'березня': '03', 'квітня': '04', 'травня': '05', 'червня': '06',
        'липня': '07', 'серпня': '08', 'вересня': '09', 'жовтня': '10', 'листопада': '11', 'грудня': '12'
    }

    date_patterns = [
        r'(\d{1,2})[./,] ?(\d{1,2})[./,] ?(\d{4})',  # Standard numeric dates
        r'(\d{1,2})[./,] ?(\d{1,2})[./,] ?(\d{2})',  # Short year numeric dates
        r'(\d{1,2})\.(\d{2})\.(\d{4})\sр',  # Dates with trailing "р"
        r'(\d{1,2}) (\w+) (\d{4})',  # Dates with month names in Ukrainian
        r'(\d{2})/(\d{2})(\d{4})',  # Dates like 08/082016 without a separator
        r'(\d{1,2})\.(\d{2}) (\d{4})',  # Dates with a space like "08.10 2020"
        r'(\d{4})(\d{2})(\d{2})'  # Incorrect concatenation like "18032019"
    ]

    for pattern in date_patterns:
        match = re.search(pattern, str(date_text))
        if match:
            groups = match.groups()
            if len(groups) == 3:
                day, month, year = groups
                if month.isdigit():
                    if len(year) == 2:
                        year = '20' + year
                    if len(day) == 1:
                        day = '0' + day
                    if len(month) == 1:
                        month = '0' + month
                else:
                    month = months_uk.get(month.lower(), month)  # Translate Ukrainian month to number
                try:
                    return datetime.strptime(f"{year}-{month}-{day}", "%Y-%m-%d")
                except ValueError as e:
                    logging.error(f"Date conversion failed for {date_text}: {str(e)}")
                    continue  # If conversion fails, try next pattern

    # Log cases where no valid date pattern matched
    logging.warning(f"No valid date pattern found for: {date_text}")
    return date_text

def get_plot(data):
    plot_number = ''
    for value in data:
        if isinstance(value, str):
            # Strip leading and trailing whitespace
            cleaned_value = value.strip()
            # Define patterns for 'Ділянка' with and without additional spaces or characters
            patterns = [
                r'Ділянка\s*№?\s*',  # Matches 'Ділянка', 'Ділянка №', 'Ділянка  №  ', etc.
                r'\s+'  # Matches any sequence of whitespace characters
            ]
            for pattern in patterns:
                cleaned_value = re.sub(pattern, '', cleaned_value).strip()
            if cleaned_value:  # Check if there is anything left worth adding
                plot_number += cleaned_value + ' '
    plot_number = plot_number.strip()
    return plot_number if plot_number else None

def get_soil_type(xls_data, idx, col_idx):
    # Initialize an empty list to store soil data from multiple lines
    combined_soil_type = []
    # Check the current row and the next two rows for soil data
    for offset in range(3):
        row_data = xls_data.iloc[idx + offset, col_idx:col_idx + 10].dropna()
        for value in row_data:
            if isinstance(value, str) and 'Ґрунт' in value:
                # Clean the soil descriptor and add it to the list
                cleaned_value = value.replace('Ґрунт', '').strip()
                combined_soil_type.append(cleaned_value)
            else:
                combined_soil_type.append(value.strip())
    # Join all collected pieces into one string
    return ' '.join(combined_soil_type)

import pandas as pd
from datetime import datetime
import re


def extract_data_from_excel(file_path):
    xls = pd.ExcelFile(file_path)
    df_final = pd.DataFrame()

    for sheet_name in xls.sheet_names:
        xls_data = pd.read_excel(xls, sheet_name=sheet_name)
        df = pd.DataFrame()
        data_index = 0
        for idx, row in xls_data.iterrows():
            for col_idx, cell in enumerate(row):
                if pd.notna(cell):
                    previous_cell = xls_data.iloc[idx - 1, col_idx] if idx > 0 else None
                    previous_3rd_cell = xls_data.iloc[idx - 3, col_idx] if idx > 2 else None

                    if "середня" in str(cell) and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                        date_slice = xls_data.iloc[max(0, idx - 3):idx + 1, 0].dropna()
                        if not date_slice.empty:
                            date = replace_date_format(date_slice.iloc[-1])
                            if date:
                                data_index = idx
                                df.loc[data_index, 'Дата'] = to_datetime(date)
                    if "Ґрунт" in str(cell):
                        soil_info = get_soil_type(xls_data, idx, col_idx)
                        df.loc[data_index, 'Ґрунт'] = soil_info
                    if "Ділянка" or "Ділянка  №           " in str(cell):
                        plot_info = [str(item) for item in xls_data.iloc[idx, col_idx + 1:col_idx + 8].dropna().tolist()]
                        df.loc[0, 'Ділянка'] = get_plot(plot_info if plot_info else [cell])
                  # Handling various conditions for moisture and density
                    conditions = {
                        "Об'ємна маса": "Об'ємна маса ґрунту на глибині ґрунту {} см, г/см3",
                        "непродуктивної": "Запаси непродуктивної вологи на глибині ґрунту {} см, мм",
                        "продуктивної при НВ": "Запаси продуктивної вологи при НВ на глибині ґрунту {} см, мм"
                    }
                    for key, format_str in conditions.items():
                        if key in str(cell):
                            data = xls_data.iloc[idx, col_idx + 1:col_idx + 51].dropna().tolist()[:10]
                            columns = [format_str.format(i * 10) for i in range(1, len(data) + 1)]
                            df.loc[data_index, columns] = data
                    # Additional moisture content conditions
                    if "середня" in str(cell) and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                        average_moisture = xls_data.iloc[idx, col_idx + 1:col_idx + 51].dropna().tolist()[
                                           :10]  # Use idx+101 to ensure covering of empty and cells with values, then [:10] took only 10 cells with values
                        average_columns = [
                            f'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту {i * 10} см, %' for i in
                            range(1, len(average_moisture) + 1)]
                        df.loc[data_index, average_columns] = average_moisture
                    # Moisture by layers
                    if "по шарах" in str(cell):
                        if "середня" in str(previous_cell):
                            total_moisture_layers = xls_data.iloc[idx, col_idx + 1:col_idx + 51].dropna().tolist()[:10]
                            total_layers_columns = [
                                f'Запаси загальної вологи по шарах на глибині ґрунту {i * 10} см, мм' for i in
                                range(1, len(total_moisture_layers) + 1)]
                            df.loc[data_index, total_layers_columns] = total_moisture_layers
                        elif "по шарах" in str(previous_cell):
                            productive_moisture_layers = xls_data.iloc[idx,
                                                         col_idx + 1:col_idx + 51].dropna().tolist()[:10]
                            productive_layers_columns = [
                                f'Запаси продуктивної вологи по шарах на глибині ґрунту {i * 10} см, мм' for i in
                                range(1, len(productive_moisture_layers) + 1)]
                            df.loc[data_index, productive_layers_columns] = productive_moisture_layers
                    if "наростаючим" in str(cell) and "по шарах" in str(previous_cell):
                        cumulative_moisture = xls_data.iloc[idx, col_idx + 1:col_idx + 51].dropna().tolist()[:10]
                        cumulative_columns = [
                            f'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту {i * 10} см, мм' for i
                            in range(1, len(cumulative_moisture) + 1)]
                        df.loc[data_index, cumulative_columns] = cumulative_moisture



            df_final = pd.concat([df_final, df], ignore_index=True)

    return df_final

path_all_folders = r"C:\Users\user\PycharmProjects\data_drought\DATA_base_soil_water"
# path_all_folders = r"C:\Users\user\PycharmProjects\data_drought\test_data"
#Create and empty database
df_DB = pd.DataFrame(columns = ['file_path', 'Рік','Область', 'Метеостанція', 'Ділянка', 'Культура', 'Ґрунт', 'Дата',
       "Об'ємна маса ґрунту на глибині ґрунту 10 см, г/см3",
       "Об'ємна маса ґрунту на глибині ґрунту 20 см, г/см3",
       "Об'ємна маса ґрунту на глибині ґрунту 30 см, г/см3",
       "Об'ємна маса ґрунту на глибині ґрунту 40 см, г/см3",
       "Об'ємна маса ґрунту на глибині ґрунту 50 см, г/см3",
       "Об'ємна маса ґрунту на глибині ґрунту 60 см, г/см3",
       "Об'ємна маса ґрунту на глибині ґрунту 70 см, г/см3",
       "Об'ємна маса ґрунту на глибині ґрунту 80 см, г/см3",
       "Об'ємна маса ґрунту на глибині ґрунту 90 см, г/см3",
       "Об'ємна маса ґрунту на глибині ґрунту 100 см, г/см3",
       'Запаси непродуктивної вологи на глибині ґрунту 10 см, мм',
       'Запаси непродуктивної вологи на глибині ґрунту 20 см, мм',
       'Запаси непродуктивної вологи на глибині ґрунту 30 см, мм',
       'Запаси непродуктивної вологи на глибині ґрунту 40 см, мм',
       'Запаси непродуктивної вологи на глибині ґрунту 50 см, мм',
       'Запаси непродуктивної вологи на глибині ґрунту 60 см, мм',
       'Запаси непродуктивної вологи на глибині ґрунту 70 см, мм',
       'Запаси непродуктивної вологи на глибині ґрунту 80 см, мм',
       'Запаси непродуктивної вологи на глибині ґрунту 90 см, мм',
       'Запаси непродуктивної вологи на глибині ґрунту 100 см, мм',
       'Запаси продуктивної вологи при НВ на глибині ґрунту 10 см, мм',
       'Запаси продуктивної вологи при НВ на глибині ґрунту 20 см, мм',
       'Запаси продуктивної вологи при НВ на глибині ґрунту 30 см, мм',
       'Запаси продуктивної вологи при НВ на глибині ґрунту 40 см, мм',
       'Запаси продуктивної вологи при НВ на глибині ґрунту 50 см, мм',
       'Запаси продуктивної вологи при НВ на глибині ґрунту 60 см, мм',
       'Запаси продуктивної вологи при НВ на глибині ґрунту 70 см, мм',
       'Запаси продуктивної вологи при НВ на глибині ґрунту 80 см, мм',
       'Запаси продуктивної вологи при НВ на глибині ґрунту 90 см, мм',
       'Запаси продуктивної вологи при НВ на глибині ґрунту 100 см, мм',
       'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 10 см, %',
       'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 20 см, %',
       'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 30 см, %',
       'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 40 см, %',
       'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 50 см, %',
       'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 60 см, %',
       'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 70 см, %',
       'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 80 см, %',
       'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 90 см, %',
       'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 100 см, %',
       'Запаси загальної вологи по шарах на глибині ґрунту 10 см, мм',
       'Запаси загальної вологи по шарах на глибині ґрунту 20 см, мм',
       'Запаси загальної вологи по шарах на глибині ґрунту 30 см, мм',
       'Запаси загальної вологи по шарах на глибині ґрунту 40 см, мм',
       'Запаси загальної вологи по шарах на глибині ґрунту 50 см, мм',
       'Запаси загальної вологи по шарах на глибині ґрунту 60 см, мм',
       'Запаси загальної вологи по шарах на глибині ґрунту 70 см, мм',
       'Запаси загальної вологи по шарах на глибині ґрунту 80 см, мм',
       'Запаси загальної вологи по шарах на глибині ґрунту 90 см, мм',
       'Запаси загальної вологи по шарах на глибині ґрунту 100 см, мм',
       'Запаси продуктивної вологи по шарах на глибині ґрунту 10 см, мм',
       'Запаси продуктивної вологи по шарах на глибині ґрунту 20 см, мм',
       'Запаси продуктивної вологи по шарах на глибині ґрунту 30 см, мм',
       'Запаси продуктивної вологи по шарах на глибині ґрунту 40 см, мм',
       'Запаси продуктивної вологи по шарах на глибині ґрунту 50 см, мм',
       'Запаси продуктивної вологи по шарах на глибині ґрунту 60 см, мм',
       'Запаси продуктивної вологи по шарах на глибині ґрунту 70 см, мм',
       'Запаси продуктивної вологи по шарах на глибині ґрунту 80 см, мм',
       'Запаси продуктивної вологи по шарах на глибині ґрунту 90 см, мм',
       'Запаси продуктивної вологи по шарах на глибині ґрунту 100 см, мм',
       'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 10 см, мм',
       'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 20 см, мм',
       'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 30 см, мм',
       'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 40 см, мм',
       'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 50 см, мм',
       'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 60 см, мм',
       'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 70 см, мм',
       'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 80 см, мм',
       'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 90 см, мм',
       'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 100 см, мм'])

import os

for year in ['2016', '2017', '2018', '2019', '2020', '2021']:
    year_path = f"{path_all_folders}/ТСГ-6_{year}"
    print(year_path)
    if os.path.exists(year_path):
        for folder_obl in os.listdir(year_path):
            folder_obl_path = os.path.join(year_path, folder_obl)
            if os.path.exists(folder_obl_path):
                for folder_st in os.listdir(folder_obl_path):
                    folder_st_path = os.path.join(folder_obl_path, folder_st)
                    if os.path.exists(folder_st_path):
                        for folder_crop in os.listdir(folder_st_path):
                            folder_crop_path = os.path.join(folder_st_path, folder_crop)
                            if os.path.exists(folder_crop_path):
                                for file in os.listdir(folder_crop_path):
                                    file_path = os.path.join(folder_crop_path, file)
                                    try:
                                        print(f"Processing file: {file_path}")
                                        df_new = extract_data_from_excel(file_path=file_path,)
                                        df_new.loc[0, 'file_path'] = file_path
                                        df_new.loc[0, 'Рік'] = year
                                        df_new.loc[0, 'Область'] = folder_obl[11:]
                                        df_new.loc[0, 'Метеостанція'] = folder_st
                                        df_new.loc[0, 'Культура'] = folder_crop
                                            # Додаємо нові дані
                                        df_DB = pd.concat([df_DB, df_new], ignore_index=True)
                                    except Exception as e:
                                        print(f"Error processing file {file_path}: {e}")
    else:
        print(f"Year directory {year_path} does not exist.")
#df_DB.to_excel('test.xlsx')
df_DB.to_excel('DB_Soil_water_vuprav2018brodu.xlsx')
