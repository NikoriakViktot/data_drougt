import os
import logging
import pandas as pd
from datetime import datetime
import re
logging.basicConfig(filename='error_log.txt', level=logging.ERROR)

# Функція для перевірки, чи є значення числом (int або float)
def is_number(x):
    return pd.to_numeric(x, errors='coerce').notnull()
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
            cleaned_value = value.strip()
            patterns = [
                r'Ділянка\s*№?\s*',  # Matches 'Ділянка', 'Ділянка №', 'Ділянка  №  ', etc.
                r'\s+'  # Matches any sequence of whitespace characters
            ]
            for pattern in patterns:
                cleaned_value = re.sub(pattern, '', cleaned_value).strip()
            if cleaned_value:
                plot_number += cleaned_value + ' '
    plot_number = plot_number.strip()
    return plot_number if plot_number else None

def extract_soil_info(cell):
    match = re.search(r'(Ґрунт|Грунт)\s*(.*)', cell)
    if match:
        return match.group(2).strip()  # Повертаємо інформацію після слова "Ґрунт" або "Грунт"
    return None

def extract_data_from_excel(file_path):
    try:
        # Визначення формату файлу та вибір правильного парсера
        if file_path.endswith('.xls'):
            xls = pd.ExcelFile(file_path, engine='xlrd')
        elif file_path.endswith('.xlsx'):
            xls = pd.ExcelFile(file_path, engine='openpyxl')
        else:
            raise ValueError(f"Unsupported file format: {file_path}")

        df_final = pd.DataFrame()

        for sheet_name in xls.sheet_names:
            try:
                xls_data = pd.read_excel(xls, sheet_name=sheet_name)
                df = pd.DataFrame()
                data_index = 0
                for idx, row in xls_data.iterrows():
                    for col_idx, cell in enumerate(row):
                        if pd.notna(cell):
                            if isinstance(cell, str):
                                previous_cell = xls_data.iloc[idx - 1, col_idx] if idx > 0 else None
                                previous_3rd_cell = xls_data.iloc[idx - 3, col_idx] if idx > 2 else None

                                if "середня" in cell and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                                    date_slice = xls_data.iloc[max(0, idx - 3):idx + 1, 0].dropna()
                                    if not date_slice.empty:
                                        date = replace_date_format(date_slice.iloc[-1])
                                        if date:
                                            data_index = idx
                                            df.loc[data_index, 'Дата'] = pd.to_datetime(date)
                                # patern_soil = re.Pattern
                                if 'Ґрунт' in cell or 'Грунт' in cell:
                                    # Використання функції для витягування інформації
                                    soil_info = extract_soil_info(cell)
                                    if soil_info:
                                        df.loc[0, 'Ґрунт'] = soil_info
                                if 'Ґрунт' in str(cell) or 'Грунт' in str(cell):
                                    soil_info = xls_data.iloc[idx, col_idx + 1:col_idx + 8].dropna().tolist()
                                    df.loc[0, 'Ґрунт'] = ' '.join(soil_info)
                                if "Ділянка" in str(cell):
                                    plot_info = [str(item) for item in
                                                 xls_data.iloc[idx, col_idx + 1:col_idx + 8].dropna().tolist()]
                                    if plot_info == []:
                                        df.loc[0, 'Ділянка'] = cell
                                    else:
                                        df.loc[0, 'Ділянка'] = ' '.join(plot_info)

                                conditions = {
                                    "Об'ємна маса": "Об'ємна маса ґрунту на глибині ґрунту {} см, г/см3",
                                    "непродуктивної": "Запаси непродуктивної вологи на глибині ґрунту {} см, мм",
                                    "продуктивної при НВ": "Запаси продуктивної вологи при НВ на глибині ґрунту {} см, мм"
                                }
                                for key, format_str in conditions.items():
                                    if key in cell:
                                        data = xls_data.iloc[idx, col_idx + 1:col_idx + 39].dropna().tolist()[:10]
                                        # Фільтруємо дані, залишаючи тільки числа
                                        # data = [d for d in data if is_number(d)]
                                        columns = [format_str.format(i * 10) for i in range(1, len(data) + 1)]
                                        df.loc[data_index, columns] = data

                                if "середня" in cell and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                                    average_moisture = xls_data.iloc[idx, col_idx + 1:col_idx + 39].dropna().tolist()[
                                                       :10]
                                    # Фільтруємо дані, залишаючи тільки числа
                                    # average_moisture = [d for d in average_moisture if is_number(d)]
                                    average_columns = [
                                        f'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту {i * 10} см, %'
                                        for
                                        i in
                                        range(1, len(average_moisture) + 1)]
                                    df.loc[data_index, average_columns] = average_moisture

                                if "по шарах" in cell:
                                    if "середня" in str(previous_cell):
                                        total_moisture_layers = xls_data.iloc[idx,
                                                                col_idx + 1:col_idx + 39].dropna().tolist()[
                                                                :10]
                                        # Фільтруємо дані, залишаючи тільки числа
                                        # total_moisture_layers = [d for d in total_moisture_layers if is_number(d)]
                                        total_layers_columns = [
                                            f'Запаси загальної вологи по шарах на глибині ґрунту {i * 10} см, мм' for i
                                            in
                                            range(1, len(total_moisture_layers) + 1)]
                                        df.loc[data_index, total_layers_columns] = total_moisture_layers
                                    elif "по шарах" in str(previous_cell):
                                        productive_moisture_layers = xls_data.iloc[idx,
                                                                     col_idx + 1:col_idx + 39].dropna().tolist()[:10]
                                        # Фільтруємо дані, залишаючи тільки числа
                                        # productive_moisture_layers = [d for d in productive_moisture_layers if is_number(d)]
                                        productive_layers_columns = [
                                            f'Запаси продуктивної вологи по шарах на глибині ґрунту {i * 10} см, мм' for
                                            i in
                                            range(1, len(productive_moisture_layers) + 1)]
                                        df.loc[data_index, productive_layers_columns] = productive_moisture_layers

                                if "наростаючим" in str(cell) or "по шарах" in str(previous_cell):
                                    cumulative_moisture = xls_data.iloc[idx,
                                                          col_idx + 1:col_idx + 39].dropna().tolist()[:10]
                                    # Фільтруємо дані, залишаючи тільки числа
                                    # cumulative_moisture = [d for d in cumulative_moisture if is_number(d)]
                                    cumulative_columns = [
                                        f'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту {i * 10} см, мм'
                                        for i in range(1, len(cumulative_moisture) + 1)]
                                    df.loc[data_index, cumulative_columns] = cumulative_moisture

                        if not df.empty:
                            df_final = pd.concat([df_final, df], ignore_index=True)

            except Exception as e:
                logging.error(f"Error reading sheet {sheet_name} in file {file_path}: {e}")
                continue
        return df_final
    except Exception as e:
        logging.error(f"Error processing file {file_path}: {e}")
        return pd.DataFrame()


path_all_folders = r"C:\Users\5302\PycharmProjects\data_drougt\DATA_base_soil_water_cleaned"
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


# for year in ['2016', '2017', '2018', '2019', '2020', '2021']:
for year in ['2016']:
    year_path = os.path.join(path_all_folders, f'ТСГ-6_{year}')
    if os.path.exists(year_path):
        for folder_obl in os.listdir(year_path):
            folder_obl_path = os.path.join(year_path, folder_obl)
            for folder_st in os.listdir(folder_obl_path):
                folder_st_path = os.path.join(folder_obl_path, folder_st)
                for folder_crop in os.listdir(folder_st_path):
                    folder_crop_path = os.path.join(folder_st_path, folder_crop)
                    for file in os.listdir(folder_crop_path):
                        file_path = os.path.join(folder_crop_path, file)
                        try:
                            print(f"Processing file: {file_path}")
                            df_new = extract_data_from_excel(file_path)
                            if not df_new.empty:
                                df_new.loc[0, 'file_path'] = file_path
                                df_new.loc[0, 'Рік'] = year
                                df_new.loc[0, 'Область'] = folder_obl[11:]
                                df_new.loc[0, 'Метеостанція'] = folder_st
                                df_new.loc[0, 'Культура'] = folder_crop.lower()
                                df_DB = pd.concat([df_DB, df_new], ignore_index=True)
                        except Exception as e:
                            logging.error(f"Error processing file {file_path}: {e}")

# Збереження результатів
output_file = 'DB_Soil_water_2016.xlsx'
df_DB.to_excel(output_file, index=False)
print(f"Processing completed. Results saved to {output_file}")