import os
import logging
import pandas as pd
from datetime import datetime
import re

logging.basicConfig(filename='error_log.txt', level=logging.ERROR)


def replace_date_format(date_text, set_day_to_one=False):
    """
    Зчитує дату із рядка.

    Параметри:
        date_text (str або datetime): текстова дата для парсингу.
        set_day_to_one (bool): якщо True, то повертається дата з днем, встановленим у 1 (для SOIL).
                               Якщо False, повертається дата як є (для MOISTURE).

    Повертає:
        datetime або None: отриману дату або None, якщо не вдалося розпізнати формат.
    """
    if isinstance(date_text, datetime):
        return date_text.replace(day=1) if set_day_to_one else date_text

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
            if not month.isdigit():
                month = months_uk.get(month.lower(), month)
            if len(year) == 2:
                year = '20' + year
            if len(day) == 1:
                day = '0' + day
            if len(month) == 1:
                month = '0' + month
            try:
                dt = datetime.strptime(f"{year}-{month}-{day}", "%Y-%m-%d")
                return dt.replace(day=1) if set_day_to_one else dt
            except ValueError:
                continue
    logging.warning(f"No valid date pattern found for: {date_text}")
    return None


def extract_soil_info(cell):
    """Витягує інформацію після слів 'Ґрунт' або 'Грунт'."""
    match = re.search(r'(Ґрунт|Грунт)\s*(.*)', cell)
    if match:
        return match.group(2).strip()
    return None

def clean_soil_string(value: str) -> str:
    """
    Очищує вхідний рядок, видаляючи зайві згадки слова 'ґрунт'/'грунт',
    розділові знаки та зайві пробіли, залишаючи тільки назву ґрунту.
    """

    # Переконуємося, що працюємо з рядком
    text = str(value).strip()

    # 1) Видаляємо всі згадки 'ґрунт' або 'грунт' (у будь-якому регістрі)
    #    \b – це межа слова, (?i) – ігнорування регістру
    text = re.sub(r'(?i)\b[ґг]рунт\b', '', text).strip()

    # 2) Видаляємо небажані розділові знаки з кінця рядка (крапки, коми, тощо)
    text = re.sub(r'[.,:;!?]+$', '', text).strip()

    # 3) Зводимо послідовні пробіли до одного та ще раз прибираємо зайві пробіли з країв
    text = re.sub(r'\s+', ' ', text).strip()

    return text


def extract_plot_number(value):
    """
    Обробляє вхідне значення або список значень, видаляючи слово "Ділянка" (без врахування регістру)
    та символ "№" (якщо він є), повертаючи лише числову частину (тобто рядок, що містить цифри).

    Якщо в результаті очищення не знайдено жодної цифри, функція повертає None.

    Приклади:
      "25"            -> "25"
      "Ділянка № 25"  -> "25"
      "Ділянка  25"   -> "25"
      "№ 25"         -> "25"

    Якщо передано список, функція перебирає елементи і повертає перший знайдений номер.
    """

    def process(text):
        # Переконуємося, що працюємо з рядком
        if not isinstance(text, str):
            text = str(text)
        # Видаляємо слово "Ділянка" з опціональним символом "№" та пробілами
        cleaned = re.sub(r'(?i)\s*Ділянка\s*№?\s*', '', text).strip()
        # Видаляємо, якщо залишився, символ "№" на початку рядка
        cleaned = re.sub(r'(?i)^№\s*', '', cleaned).strip()
        # Якщо в очищеному тексті є хоча б одна цифра, повертаємо його, інакше - None
        return cleaned if re.search(r'\d', cleaned) else None

    if isinstance(value, list):
        for item in value:
            res = process(item)
            if res is not None:
                return res
        return None
    else:
        return process(value)

class ExcelReader:
    """Відповідає за завантаження Excel-файлу та отримання даних із аркушів."""

    def __init__(self, file_path):
        self.file_path = file_path
        self.engine = None
        if file_path.endswith('.xls'):
            self.engine = 'xlrd'
        elif file_path.endswith('.xlsx'):
            self.engine = 'openpyxl'
        else:
            raise ValueError(f"Unsupported file format: {file_path}")

    def read_sheets(self):
        """Повертає словник: {sheet_name: DataFrame}."""
        try:
            xls = pd.ExcelFile(self.file_path, engine=self.engine)
            sheets = {}
            for sheet in xls.sheet_names:
                sheets[sheet] = pd.read_excel(xls, sheet_name=sheet)
            return sheets
        except Exception as e:
            logging.error(f"Error reading file {self.file_path}: {e}")
            return {}


class SoilMassParser:
    """Відповідає за парсинг даних ґрунтової маси із даних Excel."""

    def __init__(self, sheets):
        self.sheets = sheets
        # Набір "важливих" полів для перевірки даних
        self.data_fields = {
            "Об'ємна маса ґрунту на глибині ґрунту 10 см, г/см3",
            "Об'ємна маса ґрунту на глибині ґрунту 20 см, г/см3"
        }

    def parse(self):
        rows = []
        current_row = {}
        current_date = None
        for sheet_name, df in self.sheets.items():
            try:
                for idx, row in df.iterrows():
                    row_text_list = df.iloc[idx].astype(str).dropna().tolist()
                    row_text = ' '.join(row_text_list)
                    new_date_found = None
                    for col_idx, cell in enumerate(row):
                        if pd.notna(cell) and isinstance(cell, str):
                            previous_cell = df.iloc[idx - 1, col_idx] if idx > 0 else None
                            previous_3rd_cell = df.iloc[idx - 3, col_idx] if idx > 2 else None

                            # Парсинг дати (використовуємо replace_date_format_soil)
                            if "середня" in cell and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                                date_slice = df.iloc[max(0, idx - 3):idx + 1, 0].dropna()
                                if not date_slice.empty:
                                    d = replace_date_format(date_slice.iloc[0], set_day_to_one=True)
                                    if d:
                                        new_date_found = d

                            # Парсинг інформації про ґрунт
                            if 'Ґрунт' in cell or 'Грунт' in cell:
                                soil_info = extract_soil_info(cell)
                                if soil_info:
                                    current_row['Ґрунт'] = soil_info
                                soil_info_list = df.iloc[idx, col_idx + 1:col_idx + 8].dropna().tolist()
                                if soil_info_list:
                                    current_row['Ґрунт'] = ' '.join(soil_info_list)

                                    # 3) Парсинг номеру ділянки
                                    #    Додаємо умову: якщо 'Ділянка' вже є (і не None), то не перезаписуємо
                            if "Ділянка" in cell and not current_row.get('Ділянка'):
                                plot_info = [
                                            str(item)
                                            for item in df.iloc[idx, col_idx + 1:col_idx + 10].dropna().tolist()
                                        ]

                                if plot_info:
                                    current_row['Ділянка'] = extract_plot_number(plot_info)
                                else:
                                    current_row['Ділянка'] = extract_plot_number(cell)

                            # Парсинг показників (об'ємна маса, запаси вологи тощо)
                            conditions = {
                                "Об'ємна маса": "Об'ємна маса ґрунту на глибині ґрунту {} см, г/см3",
                                "непродуктивної": "Запаси непродуктивної вологи на глибині ґрунту {} см, мм",
                                "продуктивної при НВ": "Запаси продуктивної вологи при НВ на глибині ґрунту {} см, мм"
                            }
                            for key, format_str in conditions.items():
                                if key in cell:
                                    data = df.iloc[idx, col_idx + 1:col_idx + 36].dropna().tolist()[:10]
                                    columns = [format_str.format(i * 10) for i in range(1, len(data) + 1)]
                                    for i, col_name in enumerate(columns):
                                        current_row[col_name] = data[i]

                    if new_date_found:
                        if current_date is not None:
                            row_to_add = current_row.copy()
                            row_to_add['Дата'] = current_date
                            if any(field in row_to_add for field in self.data_fields):
                                rows.append(row_to_add)
                            current_row = {}
                        current_date = new_date_found

                if current_date is not None and current_row:
                    row_to_add = current_row.copy()
                    row_to_add['Дата'] = current_date
                    if any(field in row_to_add for field in self.data_fields):
                        rows.append(row_to_add)
                current_row = {}
                current_date = None

            except Exception as e:
                logging.error(f"Error reading sheet {sheet_name}: {e}")
                continue
        return pd.DataFrame(rows)



class MoistureDateParser:
    """
    Аналог "старого" підходу: проходимо по рядках DataFrame,
    шукаємо ключові слова ("середня", "по шарах" тощо),
    і зберігаємо все у df_result з індексом data_index = idx,
    коли знаходимо дату.
    """
    def __init__(self, df):
        self.df = df

    def parse_sheet_into_df(self):
        """
        Повертає DataFrame, де кожен рядок відповідає рядку Excel,
        у якому була знайдена дата. Поля (вологість, запаси вологи, ґрунт, ділянка)
        записуються у df.loc[data_index, ...].
        """
        df_result = pd.DataFrame()
        data_index = 0  # індекс рядка в df_result
        current_row = {}


        for idx, row in self.df.iterrows():
            row_text_list = self.df.iloc[idx].astype(str).dropna().tolist()
            row_text = ' '.join(row_text_list)
            for col_idx, cell in enumerate(row):
                if pd.notna(cell):
                    previous_cell = self.df.iloc[idx-1, col_idx] if idx > 0 else None
                    previous_3rd_cell = self.df.iloc[idx-3, col_idx] if idx > 2 else None

                    # 1) Шукаємо дату за ключем "середня" + (4 чи 2)
                    if "середня" in str(cell) and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                        date_slice = self.df.iloc[max(0, idx-3):idx+1, 0].dropna()
                        date_val = date_slice.iloc[-1] if not date_slice.empty else None
                        parsed_date = replace_date_format(date_val, set_day_to_one=False)
                        if parsed_date is not None:
                            data_index = idx  # Записувати дані в рядок з індексом idx
                            df_result.loc[data_index, 'Дата'] = parsed_date


                    # 2) Ґрунт (можна писати у data_index, або завжди в 0)
                    if "Ґрунт" in str(cell) or "Грунт" in str(cell):
                        soil_info_list = self.df.iloc[idx, col_idx+1:col_idx+8].dropna().tolist()
                        # СТАРИЙ код: df_result.loc[0, 'Ґрунт'] = ' '.join(map(str, soil_info_list))
                        # Якщо хочете, щоб для кожної дати ґрунт повторювався, пишемо:
                        df_result.loc[data_index, 'Ґрунт'] = ' '.join(map(str, soil_info_list))


                    # 3) Ділянка
                    if "Ділянка" in str(cell):
                        # Отримуємо дані з клітинок праворуч (від col_idx+1 до col_idx+8)
                        plot_info = [str(item) for item in self.df.iloc[idx, col_idx + 1:col_idx + 8].dropna().tolist()]

                        # Якщо список НЕ порожній, використовуємо перший знайдений елемент або об'єднуємо їх
                        if plot_info:
                            df_result.loc[data_index, 'Ділянка'] = extract_plot_number(plot_info[0])
                        else:
                            df_result.loc[data_index, 'Ділянка'] = extract_plot_number(cell)



                    # 5) "середня" вологість
                    if "середня" in str(cell) and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                        average_moisture = self.df.iloc[idx, col_idx+1:col_idx+101].dropna().tolist()[:10]
                        avg_columns = [
                            f'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту {i*10} см, %'
                            for i in range(1, len(average_moisture)+1)
                        ]
                        df_result.loc[data_index, avg_columns] = average_moisture


                    # 6) "по шарах" - загальна / продуктивна волога
                    if "по шарах" in str(cell):
                        if "середня" in str(previous_cell):
                            total_moisture_layers = self.df.iloc[idx, col_idx + 1:col_idx + 101].dropna().tolist()[:10]
                            # Замінюємо мінусові значення на 0
                            total_moisture_layers = [0 if (isinstance(x, (int, float)) and x < 0) else x for x in
                                                     total_moisture_layers]
                            total_cols = [
                                f'Запаси загальної вологи по шарах на глибині ґрунту {i * 10} см, мм'
                                for i in range(1, len(total_moisture_layers) + 1)
                            ]
                            df_result.loc[data_index, total_cols] = total_moisture_layers
                        elif "по шарах" in str(previous_cell):
                            productive_moisture_layers = self.df.iloc[idx, col_idx + 1:col_idx + 101].dropna().tolist()[
                                                         :10]
                            # Замінюємо мінусові значення на 0
                            productive_moisture_layers = [0 if (isinstance(x, (int, float)) and x < 0) else x for x in
                                                          productive_moisture_layers]
                            prod_cols = [
                                f'Запаси продуктивної вологи по шарах на глибині ґрунту {i * 10} см, мм'
                                for i in range(1, len(productive_moisture_layers) + 1)
                            ]
                            df_result.loc[data_index, prod_cols] = productive_moisture_layers

                    # 7) "наростаючим" (продуктивна волога)
                    if "наростаючим" in str(cell) and "по шарах" in str(previous_cell):
                        cumulative_moisture = self.df.iloc[idx, col_idx + 1:col_idx + 101].dropna().tolist()[:10]
                        # Замінюємо мінусові значення на 0
                        cumulative_moisture = [0 if (isinstance(x, (int, float)) and x < 0) else x for x in
                                               cumulative_moisture]
                        cum_cols = [
                            f'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту {i * 10} см, мм'
                            for i in range(1, len(cumulative_moisture) + 1)
                        ]
                        df_result.loc[data_index, cum_cols] = cumulative_moisture

        return df_result



class MoistureParser:
    """
    Приймає словник {sheet_name: DataFrame},
    викликає MoistureDateParser для кожного аркуша,
    і об'єднує все в один DataFrame.
    """
    def __init__(self, sheets):
        self.sheets = sheets

    def parse(self):
        all_dfs = []
        for sheet_name, df in self.sheets.items():
            parser = MoistureDateParser(df)
            df_moist = parser.parse_sheet_into_df()
            if not df_moist.empty:
                all_dfs.append(df_moist)
        if all_dfs:
            return pd.concat(all_dfs, ignore_index=True)
        else:
            return pd.DataFrame()


class DataBuilder:
    """Відповідає за побудову фінальної структури даних із отриманих DataFrame."""

    @staticmethod
    def build_soil_dataframe(df):
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
        if not df.empty:
            return df.reindex(columns=soil_columns)
        return df



class DataProcessor:
    """Оркеструє процес обробки файлів та побудови нової структури даних."""

    def __init__(self, base_folder, years):
        self.base_folder = base_folder
        self.years = years
        self.soil_data = pd.DataFrame()
        self.moisture_data = pd.DataFrame()

    def process(self):
        for year in self.years:
            year_path = os.path.join(self.base_folder, f'ТСГ-6_{year}')
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
                                    # Читаємо Excel-файл
                                    reader = ExcelReader(file_path)
                                    sheets = reader.read_sheets()

                                    # Парсимо дані ґрунтової маси
                                    soil_parser = SoilMassParser(sheets)
                                    df_soil = soil_parser.parse()
                                    if not df_soil.empty:
                                        df_soil['file_path'] = file_path
                                        df_soil['Рік'] = year
                                        df_soil['Область'] = folder_obl[11:]
                                        df_soil['Метеостанція'] = folder_st
                                        df_soil['Культура'] = folder_crop.lower()
                                        self.soil_data = pd.concat([self.soil_data, df_soil], ignore_index=True)

                                except Exception as e:
                                    logging.error(f"Error processing file {file_path}: {e}")

        # Фінальне форматування (видалення NaN-стовпців)
        self.soil_data.dropna(axis=1, how='all', inplace=True)

    def save(self, soil_output_file):
        soil_df = DataBuilder.build_soil_dataframe(self.soil_data)
        soil_df.to_excel(soil_output_file, index=False)
        print(f"Processing completed.\nSoil data -> {soil_output_file}\nMoisture data -> ")


# --- Основний запуск ---
if __name__ == '__main__':
    base_folder = r"C:\Users\5302\PycharmProjects\data_drougt\DATA_base_soil_water_cleaned"
    years = ['2016', '2017', '2018', '2019', '2020', '2021']
    # years = ['2017']
    processor = DataProcessor(base_folder, years)
    processor.process()
    processor.save('DB_Soil_mass_ver_1_1.xlsx')

