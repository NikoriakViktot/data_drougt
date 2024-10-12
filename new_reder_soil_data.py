import pandas as pd
from dateutil import parser
import re
from datetime import datetime
import logging
import os

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class SoilData:
    column_new_data=['file_path', 'Рік', 'Область', 'Метеостанція', 'Ділянка', 'Культура', 'Тип ґрунту', 'Дата',
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
                          'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 100 см, мм']

    def __init__(self, file):
        self.name_file = None
        self.year = None
        self.region = None
        self.weather_station = None
        self.month = None
        self.date_1 = None
        self.date_2 = None
        self.date_3 = None
        self.step_1 = None
        self.step_2 = None
        self.step_3 = None
        self.step_4 = None
        self.culture = None
        self.previous_culture = None
        self.soil_type = None
        self.plot = None
        self.data_sheets = self.open_data(file_path=file)
        # self.process_all_sheets()
        self.process_file()
        print("Accessing data_frame", self.data_frame)

        self.data_frame = pd.DataFrame(columns=self.column_new_data)

    def open_data(self, file_path):
        self.name_file = file_path
        xls = pd.ExcelFile(file_path)
        data_sheets = {}
        for sheet_name in xls.sheet_names:
            data_sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
        return data_sheets

    def process_file(self):
        xls = pd.ExcelFile(self.name_file)
        all_frames = []
        for sheet_name in xls.sheet_names:
            sheet_data = pd.read_excel(self.name_file, sheet_name=sheet_name)
            processed_data = self.read_data(sheet_data)
            if not processed_data.empty:
                all_frames.append(processed_data)
        self.data_frame = pd.concat(all_frames, ignore_index=True) if all_frames else pd.DataFrame()

    # Об'єднання всіх DataFrame зі списку

    def read_data(self, data):
        # Логіка обробки даних для кожного аркуша
        df_final = pd.DataFrame()
        steps = [self.get_first(data, start_row=12, end_row=17),
                 self.day_1(data),
                 self.day_2(data),
                 self.day_3(data)]
        for step in steps:
            if isinstance(step, pd.DataFrame) and not step.empty:
                df_final = pd.concat([df_final, step], ignore_index=True)
        return df_final

    def get_first(self, data, start_row, end_row):
        df = pd.DataFrame()

        conditions = {
            "Об'ємна маса": "Об'ємна маса ґрунту на глибині ґрунту {} см, г/см3",
            "непродуктивної": "Запаси непродуктивної вологи на глибині ґрунту {} см, мм",
            "продуктивної при НВ": "Запаси продуктивної вологи при НВ на глибині ґрунту {} см, мм"
        }
        for idx, row in data.iterrows():
            if idx >= start_row and idx <= end_row:
                for col_idx, cell in enumerate(row):
                    for key, format_str in conditions.items():
                        if key in str(cell):
                            values = data.iloc[idx, col_idx + 1:col_idx + 11].dropna().tolist()
                            columns = [format_str.format(i * 10) for i in range(1, len(values) + 1)]
                            temp_df = pd.DataFrame([values], columns=columns)
                            df = pd.concat([df, temp_df], ignore_index=True)

        # Збереження інформації про файл і контекст
        new_row = {
            'file_path': self.name_file,
            'Рік': self.extract_year_from_filename(),
            'Область': self.extract_region_from_path(),
            'Метеостанція': self.extract_station_from_path(),
            'Культура': self.extract_culture_from_path(),
            'Попередня культура': self.get_previous_culture(data),
            'Тип ґрунту': self.get_soil_type(data),
            'Ділянка': self.get_plot(data),
        }
        new_row_df = pd.DataFrame([new_row], index=[0])

        # Конкатенація нових даних з основним DataFrame
        combined_df = pd.concat([new_row_df, df], axis=1)
        return combined_df
        # self.data_frame = pd.concat([self.data_frame, combined_df], ignore_index=True)

    def get_day_soil(self, data, start_row, end_row):
        df = pd.DataFrame()
        for idx, row in data.iterrows():
            # Перевірка, чи індекс рядка належить заданому діапазону
            if idx >= start_row and idx <= end_row:
                for col_idx, cell in enumerate(row):
                    if pd.notna(cell):
                        previous_cell = data.iat[idx - 1, col_idx] if idx > 0 else None
                        previous_3rd_cell = data.iat[idx - 3, col_idx] if idx > 2 else None

                        if "середня" in str(cell) and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                            values = data.iloc[idx, col_idx + 1:col_idx + 11].dropna().tolist()
                            columns = [
                                f'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту {i * 10} см, %' for
                                i in range(1, len(values) + 1)]

                            temp_df = pd.DataFrame([values], columns=columns, index=[idx])
                            df = pd.concat([df, temp_df], axis=1)

                        if "по шарах" in str(cell):
                            values = data.iloc[idx, col_idx + 1:col_idx + 11].dropna().tolist()
                            if "середня" in str(previous_cell):
                                columns = [f'Запаси загальної вологи по шарах на глибині ґрунту {i * 10} см, мм' for i
                                           in range(1, len(values) + 1)]
                            else:
                                columns = [f'Запаси продуктивної вологи по шарах на глибині ґрунту {i * 10} см, мм' for
                                           i in range(1, len(values) + 1)]
                            temp_df = pd.DataFrame([values], columns=columns, index=[idx])
                            df = pd.concat([df, temp_df], axis=1)

                        if "наростаючим" in str(cell) and "по шарах" in str(previous_cell):
                            values = data.iloc[idx, col_idx + 1:col_idx + 11].dropna().tolist()
                            columns = [
                                f'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту {i * 10} см, мм'
                                for i in range(1, len(values) + 1)]
                            temp_df = pd.DataFrame([values], columns=columns, index=[idx])
                            df = pd.concat([df, temp_df], axis=1)
        return df
    def day_1(self, data):
        df = pd.DataFrame()
        try:
            date_1 = self.get_first_valid_date(data, 0, (17, 29), 1, (17, 29))
            step_2 = self.get_day_soil(data, start_row=20, end_row=30)

            new_row = {
                'file_path': None,
                'Рік': self.year,
                'Область': self.region,
                'Метеостанція': self.weather_station,
                'Культура': self.culture,
                'Попередня культура': self.previous_culture,
                'Тип ґрунту': self.soil_type,
                'Ділянка': self.plot,
                'Дата': date_1

            }
            if isinstance(step_2, pd.DataFrame):
                new_row_df = pd.DataFrame([new_row])  # Перетворюємо словник в DataFrame
                return pd.concat([new_row_df, step_2], ignore_index=True)
            return pd.DataFrame([new_row])
        except:
            pass

    def day_2(self, data):

        try:
            date_2 = self.get_first_valid_date(data, 0, (34, 45), 1, (35, 45))
            step_3 = self.get_day_soil(data, start_row=40, end_row=47)
            new_row = {
                'file_path': None,
                'Рік': self.year,
                'Область': self.region,
                'Метеостанція': self.weather_station,
                'Культура': self.culture,
                'Попередня культура': self.previous_culture,
                'Тип ґрунту': self.soil_type,
                'Ділянка': self.plot,
                'Дата': date_2

            }
            if isinstance(step_3, pd.DataFrame):
                new_row_df = pd.DataFrame([new_row])  # Перетворюємо словник в DataFrame
                return pd.concat([new_row_df, step_3], ignore_index=True)
            return pd.DataFrame([new_row])
        except:
            pass

    def day_3(self, data):
        df = pd.DataFrame()
        try:
            date_3 = self.get_first_valid_date(data, 0, (50, 59), 1, (51, 59))
            step_4 = self.get_day_soil(data, start_row=55, end_row=63)

            new_row = {
                'file_path': None,
                'Рік': self.year,
                'Область': self.region,
                'Метеостанція': self.weather_station,
                'Культура': self.culture,
                'Попередня культура': self.previous_culture,
                'Тип ґрунту': self.soil_type,
                'Ділянка': self.plot,
                'Дата': date_3

            }
            if isinstance(step_4, pd.DataFrame):
                new_row_df = pd.DataFrame([new_row])  # Перетворюємо словник в DataFrame
                return pd.concat([new_row_df, step_4], ignore_index=True)
            return pd.DataFrame([new_row])
        except:
            pass


    def get_year(self, data):
        for col in [0, 1]:
            density_value = data.iloc[19:71, col]
            data_subset_cleaned = density_value.dropna()
            if len(data_subset_cleaned) > 0:
                year = (str(data_subset_cleaned.iloc[0])[0:4])
                return year
        return None

    def get_month(self, data):
        for col in [0, 1]:
            density_value = data.iloc[19:71, col]
            data_subset_cleaned = density_value.dropna()
            if len(data_subset_cleaned) > 0:
                month = (str(data_subset_cleaned.iloc[0])[5:7])
                return month
        return None

    def extract_year_from_filename(self):
        # Розділяємо повний шлях до файлу на частини
        parts = self.name_file.split('\\')
        # Отримуємо назву файла
        filename = parts[-1]
        # Шукаємо всі входження чотирицифрових чисел у назві файла
        matches = re.findall(r'\d{4}', filename)
        # Якщо знайдено більше одного року, повертаємо останній
        if matches:
            return matches[-1]
        # Якщо років не знайдено, повертаємо None
        return None

    def extract_month_from_filename(self):
        # Розділяємо повний шлях до файлу на частини
        parts = self.name_file.split('\\')
        # Отримуємо назву файла
        filename = parts[-1]
        # Шукаємо всі входження двоцифрових чисел у назві файла
        matches = re.findall(r'\b\d{2}\b', filename)
        # Якщо знайдено більше одного значення, повертаємо останнє
        if matches:
            return matches[-1]
        # Якщо значень не знайдено, повертаємо None
        return None

    def get_region(self, data):
        region_data = data.iloc[4, 20:37].dropna().str.strip()
        if not region_data.empty:
            return self.extract_region(region_data.iloc[0])
        region_data = data.iloc[3, 20:37].dropna().str.strip()
        if not region_data.empty:
            return self.extract_region(region_data.iloc[0])
        return self.extract_region_from_path()

    def extract_region(self, value):
        if 'Область' in value or 'республіка' in value:
            clean_value = value.replace('Область (республіка)', '').strip()
            return clean_value
        return value.strip()

    def extract_region_from_path(self):
        parts = self.name_file.split('\\')
        region_with_prefix = parts[-4]
        region_name = region_with_prefix.split('_')[-1]
        return region_name

    def get_weather_station(self, data):
        station_data = data.iloc[4, 1:11].dropna().str.strip()
        station_name = self._process_station_data(station_data)
        if not station_name:
            station_data = data.iloc[3, 1:11].dropna().str.strip()
            station_name = self._process_station_data(station_data)

        return station_name if station_name else self.extract_station_from_path()

    def _process_station_data(self, station_data):
        station_name = ""
        for value in station_data:
            if 'Станція' in value:
                station_name += value.replace('Станція', '').strip() + " "
            else:
                station_name += value.strip() + " "
        return station_name.strip()

    def extract_station_from_path(self):
        parts = self.name_file.split('\\')
        with_prefix = parts[-3]
        station_name = with_prefix.split('_')[-1]
        return station_name


    @staticmethod
    def replace_date_format(date_text):
        date_patterns = [
            r'(\d{2})/(\d{2})/(\d{4})',  # 18/07/2016
            r'(\d{2})\.(\d{2})\.(\d{4})\.',  # 18.07.2016.
            r'(\d{2})\.(\d{2})\.(\d{4})',  # 08.07.2015
            r'(\d{2})\.(\d{2})\.(\d{2})\sр\.',  # 28.07.16 р.
            r'(\d{2})\.(\d{2})\.(\d{2})\sр',  # 27.04.16 р
            r'(\d{2})\.(\d{2})\.(\d{2})',  # 27.04.16
            r'(\d{2})\s(\d{2})\s(\d{4})',  # 18 08 2018
            r'(\d{2})-(\d{2})-(\d{4})',  # 18-08-2019
            r'(\d{4})-(\d{2})-(\d{2})',  # 0202-04-08 (intended as 2020-04-08)
            r'(\d{2})/(\d{1,2})/(\d{4})'  # 17/7/2016
        ]
        for pattern in date_patterns:
            match = re.search(pattern, date_text)
            if match:
                day, month, year = match.groups()
                if len(year) == 2:
                    year = '20' + year
                try:
                    # Convert string to datetime object
                    return datetime.strptime(f"{year}-{month}-{day}", "%Y-%m-%d")
                except ValueError:
                    continue  # If conversion fails, try next pattern
        return None  # If no pattern matches, return None

    def get_first_valid_date(self, data, primary_col, primary_row_range, secondary_col, secondary_row_range):
        logging.info("Function get_first_valid_date started")
        logging.info(f"DataFrame shape: {data.shape}")

        # Function to process date from a specific cell
        def process_date(date):
            if pd.isna(date):
                return None
            if not isinstance(date, datetime):
                date = self.replace_date_format(str(date))
            return date

        # Check in the primary column
        logging.info(f"Checking primary column: {primary_col}")
        for row_index in range(primary_row_range[0], primary_row_range[1] + 1):
            date = data.iloc[row_index, primary_col]
            processed_date = process_date(date)
            if processed_date:
                logging.info(f"Valid date found in primary column at row {row_index}: {processed_date}")
                return processed_date

        # Check in the secondary column if no valid date was found in the primary column
        logging.info(f"Checking secondary column: {secondary_col}")
        for row_index in range(secondary_row_range[0], secondary_row_range[1] + 1):
            date = data.iloc[row_index, secondary_col]
            processed_date = process_date(date)
            if processed_date:
                logging.info(f"Valid date found in secondary column at row {row_index}: {processed_date}")
                return processed_date

        logging.info("No valid dates found")
        return None


    def try_parse_date(self, date_str):
        try:
            return parser.parse(str(date_str), dayfirst=True)
        except (ValueError, TypeError):
            return None

    def get_culture(self, data):
        data_subset = data.iloc[5, 0:5].dropna()
        if not data_subset.empty:
            return self.extract_culture(data_subset.iloc[0])
        data_subset = data.iloc[6, 0:5].dropna()
        if not data_subset.empty:
            return self.extract_culture(data_subset.iloc[0])
        return self.extract_culture_from_path()

    def extract_culture_from_path(self):
        parts = self.name_file.split('\\')
        with_prefix = parts[-2]
        culture = with_prefix.split('_')[-1]
        return culture.lower()

    def extract_culture(self, value):
        if 'Культура' in value:
            return value.replace('Культура', '').strip()
        return value.strip()

    def get_previous_culture(self, data):
        for row_index in [5, 6, 7]:
            data_subset = data.iloc[row_index, 20:37].dropna()
            for item in data_subset:
                result = self.extract_previous_culture(item)
                if result:
                    return result
        return None

    def extract_previous_culture(self, value):
        if isinstance(value, str):
            if value.strip().startswith('Попередник'):
                return value.strip()[len('Попередник'):].strip()
            else:
                return value.strip()
        return None

    def get_soil_type(self, data):
        soil_data_7 = data.iloc[7, 0:10].dropna().str.strip()
        soil_data_8 = data.iloc[8, 0:10].dropna().str.strip()
        soil_data_9 = data.iloc[9, 0:10].dropna().str.strip()

        combined_soil_type = []
        for soil_data in [soil_data_7, soil_data_8, soil_data_9]:
            for value in soil_data:
                if isinstance(value, str) and 'Ґрунт' in value:
                    # Видаляємо слово "Ґрунт" і очищуємо зайві пробіли
                    cleaned_value = value.replace('Ґрунт', '').strip()
                    combined_soil_type.append(cleaned_value)
                else:
                    combined_soil_type.append(value.strip())

        return ' '.join(combined_soil_type) if combined_soil_type else None

    def get_plot(self, data):
        data_subset = data.iloc[6, 9:15]
        data_subset_cleaned = data_subset.dropna()
        plot_number = ''
        for value in data_subset_cleaned:
            if isinstance(value, str) and 'Ділянка' in value:
                cleaned_value = value.replace('Ділянка', '').strip()
                plot_number += cleaned_value + ' '
            elif isinstance(value, str):
                plot_number += value.strip() + ''
            else:
                plot_number += str(value) + ''
        plot_number = plot_number.strip()
        if plot_number:
            return plot_number
        else:
            return None

    @staticmethod
    def clean_note_value(value):
        if re.match(r'Примітка:?\s*_+', value):
            return None
        return value.replace("Примітка:", "").strip()

    def get_notes(self, data, row_index):
        data_subset = data.iloc[row_index, 0:7]
        data_subset_cleaned = data_subset.dropna()
        notes = []
        for value in data_subset_cleaned:
            if isinstance(value, str):
                cleaned_note = self.clean_note_value(value)
                if cleaned_note:
                    notes.append(cleaned_note)
        return ", ".join(notes) if notes else None
if __name__ == '__main__':
        def process_directory(directory, output_file='combined_data.xlsx'):
            master_df = pd.DataFrame()
            for root, dirs, files in os.walk(directory):
                for file in files:
                    if file.endswith(('.xls', '.xlsx')):
                        file_path = os.path.join(root, file)
                        logging.info(f"Attempting to open file: {file_path}")
                        try:
                            soil_data = SoilData(file_path)
                            result_df = soil_data.data_frame
                            if not result_df.empty:
                                master_df = pd.concat([master_df, result_df], ignore_index=True)
                            else:
                                logging.info(f"No data extracted from {file_path}")
                        except Exception as e:
                            logging.error(f"Failed to process file {file_path}: {e}")

            if not master_df.empty:
                master_df.to_excel(output_file, index=False)
                logging.info(f"Data successfully saved to {output_file}")
            else:
                logging.info("No data to save.")

        # Виклик функції для обробки каталогу
        process_directory(r"C:\Users\user\PycharmProjects\data_drought\DATA_base_soil_water", 'DB_soil_data_16.xlsx')



        # self.soil_density_10cm = data.iat[14, 19]
        # self.soil_density_20cm = data.iat[14, 21]
        # self.soil_density_30cm = data.iat[14, 23]
        # self.soil_density_40cm = data.iat[14, 25]
        # self.soil_density_50cm = data.iat[14, 27]
        # self.soil_density_60cm = data.iat[14, 29]
        # self.soil_density_70cm = data.iat[14, 31]
        # self.soil_density_80cm = data.iat[14, 33]
        # self.soil_density_90cm = data.iat[14, 35]
        # self.soil_density_100cm = data.iat[14, 37]
        # self.unproductive_moisture_10cm = data.iat[15, 19]
        # self.unproductive_moisture_20cm = data.iat[15, 21]
        # self.unproductive_moisture_30cm = data.iat[15, 23]
        # self.unproductive_moisture_40cm = data.iat[15, 25]
        # self.unproductive_moisture_50cm = data.iat[15, 27]
        # self.unproductive_moisture_60cm = data.iat[15, 29]
        # self.unproductive_moisture_70cm = data.iat[15, 31]
        # self.unproductive_moisture_80cm = data.iat[15, 33]
        # self.unproductive_moisture_90cm = data.iat[15, 35]
        # self.unproductive_moisture_100cm = data.iat[15, 37]
        # self.productive_moisture_nv_10cm = data.iat[16, 19]
        # self.productive_moisture_nv_20cm = data.iat[16, 21]
        # self.productive_moisture_nv_30cm = data.iat[16, 23]
        # self.productive_moisture_nv_40cm = data.iat[16, 25]
        # self.productive_moisture_nv_50cm = data.iat[16, 27]
        # self.productive_moisture_nv_60cm = data.iat[16, 29]
        # self.productive_moisture_nv_70cm = data.iat[16, 31]
        # self.productive_moisture_nv_80cm = data.iat[16, 33]
        # self.productive_moisture_nv_90cm = data.iat[16, 35]
        # self.productive_moisture_nv_100cm = data.iat[16, 37]
        # self.moisture_content_dry_soil_avg_10cm_1_date = data.iat[26, 19]
        # self.moisture_content_dry_soil_avg_20cm_1_date = data.iat[26, 21]
        # self.moisture_content_dry_soil_avg_30cm_1_date = data.iat[26, 23]
        # self.moisture_content_dry_soil_avg_40cm_1_date = data.iat[26, 25]
        # self.moisture_content_dry_soil_avg_50cm_1_date = data.iat[26, 27]
        # self.moisture_content_dry_soil_avg_60cm_1_date = data.iat[26, 29]
        # self.moisture_content_dry_soil_avg_70cm_1_date = data.iat[26, 31]
        # self.moisture_content_dry_soil_avg_80cm_1_date = data.iat[26, 33]
        # self.moisture_content_dry_soil_avg_90cm_1_date = data.iat[26, 35]
        # self.moisture_content_dry_soil_avg_100cm_1_date = data.iat[26, 37]
        # self.total_moisture_by_layers_10cm_1_date = data.iat[27, 19]
        # self.total_moisture_by_layers_20cm_1_date = data.iat[27, 21]
        # self.total_moisture_by_layers_30cm_1_date = data.iat[27, 23]
        # self.total_moisture_by_layers_40cm_1_date = data.iat[27, 25]
        # self.total_moisture_by_layers_50cm_1_date = data.iat[27, 27]
        # self.total_moisture_by_layers_60cm_1_date = data.iat[27, 29]
        # self.total_moisture_by_layers_70cm_1_date = data.iat[27, 31]
        # self.total_moisture_by_layers_80cm_1_date = data.iat[27, 33]
        # self.total_moisture_by_layers_90cm_1_date = data.iat[27, 35]
        # self.total_moisture_by_layers_100cm_1_date = data.iat[27, 37]
        # self.productive_moisture_10cm_1_date = data.iat[28, 19]
        # self.productive_moisture_20cm_1_date = data.iat[28, 21]
        # self.productive_moisture_30cm_1_date = data.iat[28, 23]
        # self.productive_moisture_40cm_1_date = data.iat[28, 25]
        # self.productive_moisture_50cm_1_date = data.iat[28, 27]
        # self.productive_moisture_60cm_1_date = data.iat[28, 29]
        # self.productive_moisture_70cm_1_date = data.iat[28, 31]
        # self.productive_moisture_80cm_1_date = data.iat[28, 33]
        # self.productive_moisture_90cm_1_date = data.iat[28, 35]
        # self.productive_moisture_100cm_1_date = data.iat[28, 37]
        # self.cumulative_productive_moisture_10cm_1_date = data.iat[29, 19]
        # self.cumulative_productive_moisture_20cm_1_date = data.iat[29, 21]
        # self.cumulative_productive_moisture_30cm_1_date = data.iat[29, 23]
        # self.cumulative_productive_moisture_40cm_1_date = data.iat[29, 25]
        # self.cumulative_productive_moisture_50cm_1_date = data.iat[29, 27]
        # self.cumulative_productive_moisture_60cm_1_date = data.iat[29, 29]
        # self.cumulative_productive_moisture_70cm_1_date = data.iat[29, 31]
        # self.cumulative_productive_moisture_80cm_1_date = data.iat[29, 33]
        # self.cumulative_productive_moisture_90cm_1_date = data.iat[29, 35]
        # self.cumulative_productive_moisture_100cm_1_date = data.iat[29, 37]
        # self.moisture_content_dry_soil_avg_10cm_2_date = data.iat[42, 19]
        # self.moisture_content_dry_soil_avg_20cm_2_date = data.iat[42, 21]
        # self.moisture_content_dry_soil_avg_30cm_2_date = data.iat[42, 23]
        # self.moisture_content_dry_soil_avg_40cm_2_date = data.iat[42, 25]
        # self.moisture_content_dry_soil_avg_50cm_2_date = data.iat[42, 27]
        # self.moisture_content_dry_soil_avg_60cm_2_date = data.iat[42, 29]
        # self.moisture_content_dry_soil_avg_70cm_2_date = data.iat[42, 31]
        # self.moisture_content_dry_soil_avg_80cm_2_date = data.iat[42, 33]
        # self.moisture_content_dry_soil_avg_90cm_2_date = data.iat[42, 35]
        # self.moisture_content_dry_soil_avg_100cm_2_date = data.iat[42, 37]
        # self.total_moisture_by_layers_10cm_2_date = data.iat[43, 19]
        # self.total_moisture_by_layers_20cm_2_date = data.iat[43, 21]
        # self.total_moisture_by_layers_30cm_2_date = data.iat[43, 23]
        # self.total_moisture_by_layers_40cm_2_date = data.iat[43, 25]
        # self.total_moisture_by_layers_50cm_2_date = data.iat[43, 27]
        # self.total_moisture_by_layers_60cm_2_date = data.iat[43, 29]
        # self.total_moisture_by_layers_70cm_2_date = data.iat[43, 31]
        # self.total_moisture_by_layers_80cm_2_date = data.iat[43, 33]
        # self.total_moisture_by_layers_90cm_2_date = data.iat[43, 35]
        # self.total_moisture_by_layers_100cm_2_date = data.iat[43, 37]
        # self.productive_moisture_10cm_2_date = data.iat[44, 19]
        # self.productive_moisture_20cm_2_date = data.iat[44, 21]
        # self.productive_moisture_30cm_2_date = data.iat[44, 23]
        # self.productive_moisture_40cm_2_date = data.iat[44, 25]
        # self.productive_moisture_50cm_2_date = data.iat[44, 27]
        # self.productive_moisture_60cm_2_date = data.iat[44, 29]
        # self.productive_moisture_70cm_2_date = data.iat[44, 31]
        # self.productive_moisture_80cm_2_date = data.iat[44, 33]
        # self.productive_moisture_90cm_2_date = data.iat[44, 35]
        # self.productive_moisture_100cm_2_date = data.iat[44, 37]
        # self.cumulative_productive_moisture_10cm_2_date = data.iat[45, 19]
        # self.cumulative_productive_moisture_20cm_2_date = data.iat[45, 21]
        # self.cumulative_productive_moisture_30cm_2_date = data.iat[45, 23]
        # self.cumulative_productive_moisture_40cm_2_date = data.iat[45, 25]
        # self.cumulative_productive_moisture_50cm_2_date = data.iat[45, 27]
        # self.cumulative_productive_moisture_60cm_2_date = data.iat[45, 29]
        # self.cumulative_productive_moisture_70cm_2_date = data.iat[45, 31]
        # self.cumulative_productive_moisture_80cm_2_date = data.iat[45, 33]
        # self.cumulative_productive_moisture_90cm_2_date = data.iat[45, 35]
        # self.cumulative_productive_moisture_100cm_2_date = data.iat[45, 37]
        # self.moisture_content_dry_soil_avg_10cm_3_date = data.iat[58, 19]
        # self.moisture_content_dry_soil_avg_20cm_3_date = data.iat[58, 21]
        # self.moisture_content_dry_soil_avg_30cm_3_date = data.iat[58, 23]
        # self.moisture_content_dry_soil_avg_40cm_3_date = data.iat[58, 25]
        # self.moisture_content_dry_soil_avg_50cm_3_date = data.iat[58, 27]
        # self.moisture_content_dry_soil_avg_60cm_3_date = data.iat[58, 29]
        # self.moisture_content_dry_soil_avg_70cm_3_date = data.iat[58, 31]
        # self.moisture_content_dry_soil_avg_80cm_3_date = data.iat[58, 33]
        # self.moisture_content_dry_soil_avg_90cm_3_date = data.iat[58, 35]
        # self.moisture_content_dry_soil_avg_100cm_3_date = data.iat[58, 37]
        # self.total_moisture_by_layers_10cm_3_date = data.iat[59, 19]
        # self.total_moisture_by_layers_20cm_3_date = data.iat[59, 21]
        # self.total_moisture_by_layers_30cm_3_date = data.iat[59, 23]
        # self.total_moisture_by_layers_40cm_3_date = data.iat[59, 25]
        # self.total_moisture_by_layers_50cm_3_date = data.iat[59, 27]
        # self.total_moisture_by_layers_60cm_3_date = data.iat[59, 29]
        # self.total_moisture_by_layers_70cm_3_date = data.iat[59, 31]
        # self.total_moisture_by_layers_80cm_3_date = data.iat[59, 33]
        # self.total_moisture_by_layers_90cm_3_date = data.iat[59, 35]
        # self.total_moisture_by_layers_100cm_3_date = data.iat[59, 37]
        # self.productive_moisture_10cm_3_date = data.iat[60, 19]
        # self.productive_moisture_20cm_3_date = data.iat[60, 21]
        # self.productive_moisture_30cm_3_date = data.iat[60, 23]
        # self.productive_moisture_40cm_3_date = data.iat[60, 25]
        # self.productive_moisture_50cm_3_date = data.iat[60, 27]
        # self.productive_moisture_60cm_3_date = data.iat[60, 29]
        # self.productive_moisture_70cm_3_date = data.iat[60, 31]
        # self.productive_moisture_80cm_3_date = data.iat[60, 33]
        # self.productive_moisture_90cm_3_date = data.iat[60, 35]
        # self.productive_moisture_100cm_3_date = data.iat[60, 37]
        # self.cumulative_productive_moisture_10cm_3_date = data.iat[61, 19]
        # self.cumulative_productive_moisture_20cm_3_date = data.iat[61, 21]
        # self.cumulative_productive_moisture_30cm_3_date = data.iat[61, 23]
        # self.cumulative_productive_moisture_40cm_3_date = data.iat[61, 25]
        # self.cumulative_productive_moisture_50cm_3_date = data.iat[61, 27]
        # self.cumulative_productive_moisture_60cm_3_date = data.iat[61, 29]
        # self.cumulative_productive_moisture_70cm_3_date = data.iat[61, 31]
        # self.cumulative_productive_moisture_80cm_3_date = data.iat[61, 33]
        # self.cumulative_productive_moisture_90cm_3_date = data.iat[61, 35]
        # self.cumulative_productive_moisture_100cm_3_date = data.iat[61, 37]
        # self.ave_temperature_1_date = data.iat[68, 13]
        # self.ave_temperature_2_date = data.iat[68, 25]
        # self.ave_temperature_3_date = data.iat[68, 37]
        # self.precipitation_1_date = data.iat[70, 13]
        # self.precipitation_2_date = data.iat[70, 25]
        # self.precipitation_3_date = data.iat[70, 37]


        # self.soil_density_10cm = None
        # self.soil_density_20cm = None
        # self.soil_density_30cm = None
        # self.soil_density_40cm = None
        # self.soil_density_50cm = None
        # self.soil_density_60cm = None
        # self.soil_density_70cm = None
        # self.soil_density_80cm = None
        # self.soil_density_90cm = None
        # self.soil_density_100cm = None
        # self.unproductive_moisture_10cm = None
        # self.unproductive_moisture_20cm = None
        # self.unproductive_moisture_30cm = None
        # self.unproductive_moisture_40cm = None
        # self.unproductive_moisture_50cm = None
        # self.unproductive_moisture_60cm = None
        # self.unproductive_moisture_70cm = None
        # self.unproductive_moisture_80cm = None
        # self.unproductive_moisture_90cm = None
        # self.unproductive_moisture_100cm = None
        # self.productive_moisture_nv_10cm = None
        # self.productive_moisture_nv_20cm = None
        # self.productive_moisture_nv_30cm = None
        # self.productive_moisture_nv_40cm = None
        # self.productive_moisture_nv_50cm = None
        # self.productive_moisture_nv_60cm = None
        # self.productive_moisture_nv_70cm = None
        # self.productive_moisture_nv_80cm = None
        # self.productive_moisture_nv_90cm = None
        # self.productive_moisture_nv_100cm = None
        # self.moisture_content_dry_soil_avg_10cm_1_date = None
        # self.moisture_content_dry_soil_avg_20cm_1_date = None
        # self.moisture_content_dry_soil_avg_30cm_1_date = None
        # self.moisture_content_dry_soil_avg_40cm_1_date = None
        # self.moisture_content_dry_soil_avg_50cm_1_date = None
        # self.moisture_content_dry_soil_avg_60cm_1_date = None
        # self.moisture_content_dry_soil_avg_70cm_1_date = None
        # self.moisture_content_dry_soil_avg_80cm_1_date = None
        # self.moisture_content_dry_soil_avg_90cm_1_date = None
        # self.moisture_content_dry_soil_avg_100cm_1_date = None
        # self.total_moisture_by_layers_10cm_1_date = None
        # self.total_moisture_by_layers_20cm_1_date = None
        # self.total_moisture_by_layers_30cm_1_date = None
        # self.total_moisture_by_layers_40cm_1_date = None
        # self.total_moisture_by_layers_50cm_1_date = None
        # self.total_moisture_by_layers_60cm_1_date = None
        # self.total_moisture_by_layers_70cm_1_date = None
        # self.total_moisture_by_layers_80cm_1_date = None
        # self.total_moisture_by_layers_90cm_1_date = None
        # self.total_moisture_by_layers_100cm_1_date = None
        # self.productive_moisture_10cm_1_date = None
        # self.productive_moisture_20cm_1_date = None
        # self.productive_moisture_30cm_1_date = None
        # self.productive_moisture_40cm_1_date = None
        # self.productive_moisture_50cm_1_date = None
        # self.productive_moisture_60cm_1_date = None
        # self.productive_moisture_70cm_1_date = None
        # self.productive_moisture_80cm_1_date = None
        # self.productive_moisture_90cm_1_date = None
        # self.productive_moisture_100cm_1_date = None
        # self.cumulative_productive_moisture_10cm_1_date = None
        # self.cumulative_productive_moisture_20cm_1_date= None
        # self.cumulative_productive_moisture_30cm_1_date = None
        # self.cumulative_productive_moisture_40cm_1_date = None
        # self.cumulative_productive_moisture_50cm_1_date = None
        # self.cumulative_productive_moisture_60cm_1_date = None
        # self.cumulative_productive_moisture_70cm_1_date = None
        # self.cumulative_productive_moisture_80cm_1_date = None
        # self.cumulative_productive_moisture_90cm_1_date = None
        # self.cumulative_productive_moisture_100cm_1_date = None
        # self.moisture_content_dry_soil_avg_10cm_2_date = None
        # self.moisture_content_dry_soil_avg_20cm_2_date = None
        # self.moisture_content_dry_soil_avg_30cm_2_date = None
        # self.moisture_content_dry_soil_avg_40cm_2_date = None
        # self.moisture_content_dry_soil_avg_50cm_2_date = None
        # self.moisture_content_dry_soil_avg_60cm_2_date = None
        # self.moisture_content_dry_soil_avg_70cm_2_date = None
        # self.moisture_content_dry_soil_avg_80cm_2_date = None
        # self.moisture_content_dry_soil_avg_90cm_2_date = None
        # self.moisture_content_dry_soil_avg_100cm_2_date = None
        # self.total_moisture_by_layers_10cm_2_date = None
        # self.total_moisture_by_layers_20cm_2_date = None
        # self.total_moisture_by_layers_30cm_2_date = None
        # self.total_moisture_by_layers_40cm_2_date = None
        # self.total_moisture_by_layers_50cm_2_date = None
        # self.total_moisture_by_layers_60cm_2_date = None
        # self.total_moisture_by_layers_70cm_2_date = None
        # self.total_moisture_by_layers_80cm_2_date = None
        # self.total_moisture_by_layers_90cm_2_date = None
        # self.total_moisture_by_layers_100cm_2_date = None
        # self.productive_moisture_10cm_2_date = None
        # self.productive_moisture_20cm_2_date = None
        # self.productive_moisture_30cm_2_date = None
        # self.productive_moisture_40cm_2_date = None
        # self.productive_moisture_50cm_2_date = None
        # self.productive_moisture_60cm_2_date = None
        # self.productive_moisture_70cm_2_date = None
        # self.productive_moisture_80cm_2_date = None
        # self.productive_moisture_90cm_2_date = None
        # self.productive_moisture_100cm_2_date = None
        # self.cumulative_productive_moisture_10cm_2_date = None
        # self.cumulative_productive_moisture_20cm_2_date = None
        # self.cumulative_productive_moisture_30cm_2_date = None
        # self.cumulative_productive_moisture_40cm_2_date = None
        # self.cumulative_productive_moisture_50cm_2_date = None
        # self.cumulative_productive_moisture_60cm_2_date = None
        # self.cumulative_productive_moisture_70cm_2_date = None
        # self.cumulative_productive_moisture_80cm_2_date = None
        # self.cumulative_productive_moisture_90cm_2_date = None
        # self.cumulative_productive_moisture_100cm_2_date = None
        # self.moisture_content_dry_soil_avg_10cm_3_date = None
        # self.moisture_content_dry_soil_avg_20cm_3_date = None
        # self.moisture_content_dry_soil_avg_30cm_3_date = None
        # self.moisture_content_dry_soil_avg_40cm_3_date = None
        # self.moisture_content_dry_soil_avg_50cm_3_date = None
        # self.moisture_content_dry_soil_avg_60cm_3_date = None
        # self.moisture_content_dry_soil_avg_70cm_3_date = None
        # self.moisture_content_dry_soil_avg_80cm_3_date = None
        # self.moisture_content_dry_soil_avg_90cm_3_date = None
        # self.moisture_content_dry_soil_avg_100cm_3_date = None
        # self.total_moisture_by_layers_10cm_3_date = None
        # self.total_moisture_by_layers_20cm_3_date = None
        # self.total_moisture_by_layers_30cm_3_date = None
        # self.total_moisture_by_layers_40cm_3_date = None
        # self.total_moisture_by_layers_50cm_3_date = None
        # self.total_moisture_by_layers_60cm_3_date = None
        # self.total_moisture_by_layers_70cm_3_date = None
        # self.total_moisture_by_layers_80cm_3_date = None
        # self.total_moisture_by_layers_90cm_3_date = None
        # self.total_moisture_by_layers_100cm_3_date = None
        # self.cumulative_productive_moisture_10cm_3_date = None
        # self.productive_moisture_10cm_3_date = None
        # self.productive_moisture_20cm_3_date = None
        # self.productive_moisture_30cm_3_date = None
        # self.productive_moisture_40cm_3_date = None
        # self.productive_moisture_50cm_3_date = None
        # self.productive_moisture_60cm_3_date = None
        # self.productive_moisture_70cm_3_date = None
        # self.productive_moisture_80cm_3_date = None
        # self.productive_moisture_90cm_3_date = None
        # self.productive_moisture_100cm_3_date = None
        # self.cumulative_productive_moisture_20cm_3_date = None
        # self.cumulative_productive_moisture_30cm_3_date = None
        # self.cumulative_productive_moisture_40cm_3_date = None
        # self.cumulative_productive_moisture_50cm_3_date = None
        # self.cumulative_productive_moisture_60cm_3_date = None
        # self.cumulative_productive_moisture_70cm_3_date = None
        # self.cumulative_productive_moisture_80cm_3_date = None
        # self.cumulative_productive_moisture_90cm_3_date = None
        # self.cumulative_productive_moisture_100cm_3_date = None
        # self.ave_temperature_1_date = None
        # self.ave_temperature_2_date = None
        # self.ave_temperature_3_date = None
        # self.precipitation_1_date = None
        # self.precipitation_2_date  = None
        # self.precipitation_3_date =  None
        # self.note_1_date = None
        # self.note_2_date = None
        # self.note_3_date = None

        # def get_date_from_row(self, data, primary_col, primary_row_range, secondary_col, secondary_row_range):
    #     # Спроба знайти дату в першій колонці в заданому діапазоні рядків
    #     for row_index in range(primary_row_range[0], primary_row_range[1] + 1):
    #         # date = self.try_parse_date(data.iloc[row_index, primary_col])
    #         date = data.iloc[row_index, primary_col]
    #
    #         if date:
    #             return date
    #
    #     # Якщо дата не знайдена в першій колонці, перевіряємо другу колонку в заданому діапазоні
    #     for row_index in range(secondary_row_range[0], secondary_row_range[1] + 1):
    #         # date = self.try_parse_date(data.iloc[row_index, secondary_col])
    #         date = data.iloc[row_index, secondary_col]
    #         if date:
    #             return date
    #     return None

    # def get_date_from_row(self, data, primary_col, primary_row_range, secondary_col, secondary_row_range):
    #     # Перевірка в першій колонці
    #     for row_index in range(primary_row_range[0], primary_row_range[1] + 1):
    #         date = data.iloc[row_index, primary_col]
    #         if pd.notna(date):
    #             return date
    #
    #     # Якщо дата не знайдена в першій колонці, перевіряємо другу колонку в заданому діапазоні
    #     for row_index in range(secondary_row_range[0], secondary_row_range[1] + 1):
    #         date = data.iloc[row_index, secondary_col]
    #         if pd.notna(date):
    #             return date
    #
    #     return None