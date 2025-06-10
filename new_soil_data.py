import os
import logging
import pandas as pd
from datetime import datetime
import re

logging.basicConfig(filename='error_log.txt', level=logging.ERROR)


def replace_date_format_soil(date_text):
    """
    Функція для зчитування дати, але ВСТАНОВЛЮЄМО день = 1 (для SOIL).
    """
    if isinstance(date_text, datetime):
        return date_text.replace(day=1)

    date_text = str(date_text).strip()
    months_uk = {
        'січня': '01', 'лютого': '02', 'березня': '03', 'квітня': '04',
        'травня': '05', 'червня': '06', 'липня': '07', 'серпня': '08',
        'вересня': '09', 'жовтня': '10', 'листопада': '11', 'грудня': '12'
    }

    date_patterns = [
        r'(\d{1,2})[./, ]+(\d{1,2})[./, ]+(\d{4})',
        r'(\d{1,2})[./, ]+(\d{1,2})[./, ]+(\d{2})',
        r'(\d{1,2})\.(\d{2})\.(\d{4})\sр',
        r'(\d{1,2}) (\w+) (\d{4})',
        r'(\d{2})/(\d{2})(\d{4})',
        r'(\d{1,2})\.(\d{2}) (\d{4})',
        r'(\d{4})(\d{2})(\d{2})'
    ]

    for pattern in date_patterns:
        match = re.search(pattern, date_text)
        if match:
            day, month, year = match.groups()
            # Якщо month – слово
            if not month.isdigit():
                month = months_uk.get(month.lower(), month)

            # Якщо рік – 2 цифри
            if len(year) == 2:
                year = '20' + year

            # Додаємо нуль, якщо день/місяць складаються з 1 цифри
            if len(day) == 1:
                day = '0' + day
            if len(month) == 1:
                month = '0' + month

            try:
                dt = datetime.strptime(f"{year}-{month}-{day}", "%Y-%m-%d")
                # Ось тут змінюємо день на 1
                return dt.replace(day=1)
            except ValueError:
                continue

    logging.warning(f"No valid date pattern found for: {date_text}")
    return None


def replace_date_format_moisture(date_text):
    """
    Функція для зчитування дати, БЕЗ встановлення її на 1 число (для MOISTURE).
    """
    if isinstance(date_text, datetime):
        return date_text  # Повертаємо як є

    date_text = str(date_text).strip()
    months_uk = {
        'січня': '01', 'лютого': '02', 'березня': '03', 'квітня': '04',
        'травня': '05', 'червня': '06', 'липня': '07', 'серпня': '08',
        'вересня': '09', 'жовтня': '10', 'листопада': '11', 'грудня': '12'
    }

    date_patterns = [
        r'(\d{1,2})[./, ]+(\d{1,2})[./, ]+(\d{4})',
        r'(\d{1,2})[./, ]+(\d{1,2})[./, ]+(\d{2})',
        r'(\d{1,2})\.(\d{2})\.(\d{4})\sр',
        r'(\d{1,2}) (\w+) (\d{4})',
        r'(\d{2})/(\d{2})(\d{4})',
        r'(\d{1,2})\.(\d{2}) (\d{4})',
        r'(\d{4})(\d{2})(\d{2})'
    ]

    for pattern in date_patterns:
        match = re.search(pattern, date_text)
        if match:
            day, month, year = match.groups()
            # Якщо month – слово
            if not month.isdigit():
                month = months_uk.get(month.lower(), month)

            # Якщо рік – 2 цифри
            if len(year) == 2:
                year = '20' + year

            # Додаємо нуль, якщо день/місяць складаються з 1 цифри
            if len(day) == 1:
                day = '0' + day
            if len(month) == 1:
                month = '0' + month

            try:
                dt = datetime.strptime(f"{year}-{month}-{day}", "%Y-%m-%d")
                return dt  # Повертаємо оригінальний день
            except ValueError:
                continue

    logging.warning(f"No valid date pattern found for: {date_text}")
    return None

def extract_soil_info(cell):
    """Зчитати інформацію після слова 'Ґрунт'/'Грунт'."""
    match = re.search(r'(Ґрунт|Грунт)\s*(.*)', cell)
    if match:
        return match.group(2).strip()
    return None
def extract_soil_mass_data(file_path):
    try:
        if file_path.endswith('.xls'):
            xls = pd.ExcelFile(file_path, engine='xlrd')
        elif file_path.endswith('.xlsx'):
            xls = pd.ExcelFile(file_path, engine='openpyxl')
        else:
            raise ValueError(f"Unsupported file format: {file_path}")

        data_fields = {
            "Об'ємна маса ґрунту на глибині ґрунту 10 см, г/см3",
            "Об'ємна маса ґрунту на глибині ґрунту 20 см, г/см3"
               }

        rows = []
        current_row = {}
        current_date = None

        for sheet_name in xls.sheet_names:
            try:
                xls_data = pd.read_excel(xls, sheet_name=sheet_name)

                for idx, row in xls_data.iterrows():
                    new_date_found = None
                    for col_idx, cell in enumerate(row):
                        if pd.notna(cell) and isinstance(cell, str):
                            previous_cell = xls_data.iloc[idx - 1, col_idx] if idx > 0 else None
                            previous_3rd_cell = xls_data.iloc[idx - 3, col_idx] if idx > 2 else None

                            # Якщо "середня" і поруч "4"/"2" -> вважаємо це індикатором дати
                            if "середня" in cell and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                                date_slice = xls_data.iloc[max(0, idx - 3):idx + 1, 0].dropna()
                                if not date_slice.empty:
                                    d = replace_date_format_soil(date_slice.iloc[0])
                                    if d:
                                        new_date_found = d

                            # Ґрунт
                            if 'Ґрунт' in cell or 'Грунт' in cell:
                                soil_info = extract_soil_info(cell)
                                if soil_info:
                                    current_row['Ґрунт'] = soil_info
                                soil_info_list = xls_data.iloc[idx, col_idx + 1:col_idx + 8].dropna().tolist()
                                if soil_info_list:
                                    current_row['Ґрунт'] = ' '.join(soil_info_list)

                            # Ділянка
                            if "Ділянка" in cell:
                                plot_info = [
                                    str(item) for item in xls_data.iloc[idx, col_idx + 1:col_idx + 8].dropna().tolist()
                                ]
                                if plot_info:
                                    current_row['Ділянка'] = ' '.join(plot_info)
                                else:
                                    current_row['Ділянка'] = cell

                            # Збір даних для об'ємної маси, непродуктивної/продуктивної вологи
                            conditions = {
                                "Об'ємна маса": "Об'ємна маса ґрунту на глибині ґрунту {} см, г/см3",
                                "непродуктивної": "Запаси непродуктивної вологи на глибині ґрунту {} см, мм",
                                "продуктивної при НВ": "Запаси продуктивної вологи при НВ на глибині ґрунту {} см, мм"
                            }
                            for key, format_str in conditions.items():
                                if key in cell:
                                    data = xls_data.iloc[idx, col_idx + 1:col_idx + 36].dropna().tolist()[:10]
                                    columns = [format_str.format(i * 10) for i in range(1, len(data) + 1)]
                                    for i, col_name in enumerate(columns):
                                        current_row[col_name] = data[i]

                    # Якщо знайшли нову дату -> зберігаємо попередню, якщо є хоч якісь дані
                    if new_date_found:
                        if current_date is not None:
                            row_to_add = current_row.copy()
                            row_to_add['Дата'] = current_date
                            # Перевіряємо, чи є в row_to_add хоча б одне поле з data_fields
                            if any(field in row_to_add for field in data_fields):
                                rows.append(row_to_add)
                            current_row = {}
                        current_date = new_date_found

                # Після закінчення листа, зберігаємо поточний блок, якщо він не порожній
                if current_date is not None and current_row:
                    row_to_add = current_row.copy()
                    row_to_add['Дата'] = current_date
                    # Знову перевірка
                    if any(field in row_to_add for field in data_fields):
                        rows.append(row_to_add)

                current_row = {}
                current_date = None

            except Exception as e:
                logging.error(f"Error reading sheet {sheet_name} in file {file_path}: {e}")
                continue

        df_final = pd.DataFrame(rows)

        # Перелік усіх потенційних колонок, які вас цікавлять
        soil_columns = [
            'file_path', 'Рік', 'Область', 'Метеостанція', 'Ділянка', 'Культура', 'Ґрунт', 'Дата',
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
            'Запаси продуктивної вологи при НВ на глибині ґрунту 100 см, мм'
        ]

        if not df_final.empty:
            df_final = df_final.reindex(columns=soil_columns)

        return df_final

    except Exception as e:
        logging.error(f"Error processing file {file_path}: {e}")
        return pd.DataFrame()

def extract_moisture_data_single_row_per_date(file_path):
    try:
        if file_path.endswith('.xls'):
            xls = pd.ExcelFile(file_path, engine='xlrd')
        elif file_path.endswith('.xlsx'):
            xls = pd.ExcelFile(file_path, engine='openpyxl')
        else:
            raise ValueError(f"Unsupported file format: {file_path}")

        rows = []  # Тут складатимемо готові рядки
        current_row = {}  # Накопичуємо дані для поточної дати
        current_date = None  # Поточна дата
        print('Дані для -------', current_row)
        print('ДAТА -------',current_date)
        for sheet_name in xls.sheet_names:
            try:
                xls_data = pd.read_excel(xls, sheet_name=sheet_name)

                # Проходимо по рядках листа
                for idx, row in xls_data.iterrows():
                    # Для зручності проходимо по кожній клітинці
                    new_date_found = None
                    for col_idx, cell in enumerate(row):
                        if pd.notna(cell) and isinstance(cell, str):
                            # Шукаємо дату
                            previous_cell = xls_data.iloc[idx - 1, col_idx] if idx > 0 else None
                            previous_3rd_cell = xls_data.iloc[idx - 3, col_idx] if idx > 2 else None

                            # Якщо зустріли ключове слово "середня" + "4" / "2" - вважаємо, що це вказівка на дату
                            if "середня" in cell and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                                date_slice = xls_data.iloc[max(0, idx - 3):idx + 1, 0].dropna()
                                if not date_slice.empty:
                                    # Припустимо, беремо лише першу дату (якщо кілька)
                                    d = replace_date_format_moisture(date_slice.iloc[0])
                                    if d:
                                        new_date_found = d

                            # Шукаємо "Ґрунт"
                            if 'Ґрунт' in cell or 'Грунт' in cell:
                                soil_info = extract_soil_info(cell)
                                if soil_info:
                                    current_row['Ґрунт'] = soil_info
                                soil_info_list = xls_data.iloc[idx, col_idx + 1:col_idx + 8].dropna().tolist()
                                if soil_info_list:
                                    current_row['Ґрунт'] = ' '.join(soil_info_list)

                            # Шукаємо "Ділянка"
                            if "Ділянка" in cell:
                                plot_info = [str(item) for item in xls_data.iloc[idx, col_idx + 1:col_idx + 8].dropna().tolist()]
                                if plot_info:
                                    current_row['Ділянка'] = ' '.join(plot_info)
                                else:
                                    current_row['Ділянка'] = cell

                            # Вологість (середня)
                            if "середня" in cell and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                                average_moisture = xls_data.iloc[idx, col_idx + 1:col_idx + 39].dropna().tolist()[:10]
                                average_columns = [
                                    f'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту {i * 10} см, %'
                                    for i in range(1, len(average_moisture) + 1)
                                ]
                                for i, col_name in enumerate(average_columns):
                                    current_row[col_name] = average_moisture[i]

                            # Запаси загальної вологи по шарах
                            if "по шарах" in cell:
                                if "середня" in str(previous_cell):
                                    total_moisture_layers = xls_data.iloc[idx, col_idx + 1:col_idx + 39].dropna().tolist()[:10]
                                    total_layers_columns = [
                                        f'Запаси загальної вологи по шарах на глибині ґрунту {i * 10} см, мм'
                                        for i in range(1, len(total_moisture_layers) + 1)
                                    ]
                                    for i, col_name in enumerate(total_layers_columns):
                                        current_row[col_name] = total_moisture_layers[i]
                                elif "по шарах" in str(previous_cell):
                                    productive_moisture_layers = xls_data.iloc[idx, col_idx + 1:col_idx + 39].dropna().tolist()[:10]
                                    productive_layers_columns = [
                                        f'Запаси продуктивної вологи по шарах на глибині ґрунту {i * 10} см, мм'
                                        for i in range(1, len(productive_moisture_layers) + 1)
                                    ]
                                    for i, col_name in enumerate(productive_layers_columns):
                                        current_row[col_name] = productive_moisture_layers[i]

                            # Запаси продуктивної вологи наростаючим підсумком
                            if "наростаючим" in cell or "по шарах" in str(previous_cell):
                                cumulative_moisture = xls_data.iloc[idx, col_idx + 1:col_idx + 39].dropna().tolist()[:10]
                                cumulative_columns = [
                                    f'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту {i * 10} см, мм'
                                    for i in range(1, len(cumulative_moisture) + 1)
                                ]
                                for i, col_name in enumerate(cumulative_columns):
                                    current_row[col_name] = cumulative_moisture[i]

                    # ====== Після обробки КОЖНОГО рядка перевіряємо, чи знайшли нову дату ======
                    if new_date_found:
                        if current_date is not None:
                            # Закриваємо стару дату
                            row_to_add = current_row.copy()
                            row_to_add['Дата'] = current_date
                            rows.append(row_to_add)

                            # ВАЖЛИВО: Залишаємо 'Ґрунт' і 'Ділянка', якщо хочемо, щоб вони переходили
                            grunt = current_row.get('Ґрунт')
                            dilyanka = current_row.get('Ділянка')

                            current_row = {}
                            if grunt:
                                current_row['Ґрунт'] = grunt
                            if dilyanka:
                                current_row['Ділянка'] = dilyanka

                        current_date = new_date_found


                # Після закінчення листа, якщо current_date ще є, додаємо останній блок
                if current_date is not None and current_row:
                    row_to_add = current_row.copy()
                    row_to_add['Дата'] = current_date
                    rows.append(row_to_add)

                # Обнулити для наступного листа (якщо листів кілька)
                current_row = {}
                current_date = None

            except Exception as e:
                logging.error(f"Error reading sheet {sheet_name} in file {file_path}: {e}")
                continue

        # Створюємо фінальний DataFrame
        df_final = pd.DataFrame(rows)

        # Тепер фільтруємо потрібні колонки
        moisture_columns = [
            'file_path', 'Рік', 'Область', 'Метеостанція', 'Ділянка', 'Культура', 'Ґрунт', 'Дата',
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
            'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 100 см, мм'
        ]

        if not df_final.empty:
            df_final = df_final.reindex(columns=moisture_columns)

        return df_final

    except Exception as e:
        logging.error(f"Error processing file {file_path}: {e}")
        return pd.DataFrame()


def process_files_and_save():
    path_all_folders = r"C:\Users\5302\PycharmProjects\data_drougt\DATA_base_soil_water_cleaned"

    df_soil = pd.DataFrame()
    df_moisture = pd.DataFrame()

    for year in ['2021']:  # можна додати інші роки, якщо потрібно
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

                                # 1) Обробка даних ґрунтової маси (Soil)
                                df_new_soil = extract_soil_mass_data(file_path)
                                if not df_new_soil.empty:
                                    # Призначення даних з назв папок для кожного рядка DataFrame
                                    df_new_soil['file_path'] = file_path
                                    df_new_soil['Рік'] = year
                                    df_new_soil['Область'] = folder_obl[11:]  # наприклад, "Івано-Франківська"
                                    df_new_soil['Метеостанція'] = folder_st
                                    df_new_soil['Культура'] = folder_crop.lower()

                                    df_soil = pd.concat([df_soil, df_new_soil], ignore_index=True)

                                # 2) Обробка даних по вологості (Moisture)
                                df_new_moist = extract_moisture_data_single_row_per_date(file_path)
                                if not df_new_moist.empty:
                                    # Призначення метаданих для кожного рядка
                                    df_new_moist['file_path'] = file_path
                                    df_new_moist['Рік'] = year
                                    df_new_moist['Область'] = folder_obl[11:]
                                    df_new_moist['Метеостанція'] = folder_st
                                    df_new_moist['Культура'] = folder_crop.lower()

                                    df_moisture = pd.concat([df_moisture, df_new_moist], ignore_index=True)

                            except Exception as e:
                                logging.error(f"Error processing file {file_path}: {e}")

    # Видаляємо стовпці, де всі значення – NaN, щоб уникнути FutureWarning
    df_soil.dropna(axis=1, how='all', inplace=True)
    df_moisture.dropna(axis=1, how='all', inplace=True)

    soil_output_file = 'DB_Soil_mass_2021.xlsx'
    moisture_output_file = 'DB_Moisture_2021.xlsx'

    df_soil.to_excel(soil_output_file, index=False)
    df_moisture.to_excel(moisture_output_file, index=False)

    print(f"Processing completed.\nSoil data -> {soil_output_file}\nMoisture data -> {moisture_output_file}")

process_files_and_save()