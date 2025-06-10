
import os
import logging
import pandas as pd
from datetime import datetime
import re

# Вмикаємо відладочний режим
DEBUG = True


def debug_print(*args):
    if DEBUG:
        print(*args)


logging.basicConfig(filename='error_log.txt', level=logging.ERROR)


def replace_date_format(date_text, set_day_to_one=False):
    """
    Парсить дату із рядка і виводить інформацію для відладки.
    """
    debug_print("replace_date_format - input type:", type(date_text), "value:", date_text)
    if isinstance(date_text, datetime):
        result = date_text.replace(day=1) if set_day_to_one else date_text
        debug_print("replace_date_format - datetime input, result:", result)
        return result

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
            debug_print(f"replace_date_format - matched pattern: {pattern} -> groups:", match.groups())
            if not month.isdigit():
                month = months_uk.get(month.lower(), month)
                debug_print("replace_date_format - converted month:", month)
            if len(year) == 2:
                year = '20' + year
                debug_print("replace_date_format - converted 2-digit year:", year)
            if len(day) == 1:
                day = '0' + day
            if len(month) == 1:
                month = '0' + month
            try:
                dt = datetime.strptime(f"{year}-{month}-{day}", "%Y-%m-%d")
                result = dt.replace(day=1) if set_day_to_one else dt
                debug_print("replace_date_format - parsed date:", result)
                return result
            except ValueError as ve:
                debug_print("replace_date_format - ValueError:", ve)
                continue
    logging.warning(f"No valid date pattern found for: {date_text}")
    debug_print("replace_date_format - no valid date found for:", date_text)
    return None


def extract_soil_info(cell):
    """Витягує інформацію після слів 'Ґрунт' або 'Грунт' із рядка."""
    debug_print("extract_soil_info - input cell:", cell, "type:", type(cell))
    match = re.search(r'(Ґрунт|Грунт)\s*(.*)', cell)
    if match:
        result = match.group(2).strip()
        debug_print("extract_soil_info - extracted:", result)
        return result
    debug_print("extract_soil_info - no match found in cell.")
    return None


def extract_plot_number(value):
    """
    Видаляє слово "Ділянка" та символ "№" із тексту (чи списку текстів)
    та повертає знайдену числову частину. Виводить дані для відладки.
    """
    debug_print("extract_plot_number - input value:", value, "type:", type(value))

    def process(text):
        if not isinstance(text, str):
            text = str(text)
        debug_print("extract_plot_number.process - processing text:", text)
        cleaned = re.sub(r'(?i)\s*Ділянка\s*№?\s*', '', text).strip()
        cleaned = re.sub(r'(?i)^№\s*', '', cleaned).strip()
        debug_print("extract_plot_number.process - cleaned text:", cleaned)
        return cleaned if re.search(r'\d', cleaned) else None

    if isinstance(value, list):
        for item in value:
            res = process(item)
            if res is not None:
                debug_print("extract_plot_number - result from list:", res)
                return res
        debug_print("extract_plot_number - no digits found in list.")
        return None
    else:
        result = process(value)
        debug_print("extract_plot_number - result:", result)
        return result


class ExcelReader:
    """Читає Excel‑файл та повертає дані з аркушів."""

    def __init__(self, file_path):
        self.file_path = file_path
        self.engine = None
        if file_path.endswith('.xls'):
            self.engine = 'xlrd'
        elif file_path.endswith('.xlsx'):
            self.engine = 'openpyxl'
        else:
            raise ValueError(f"Unsupported file format: {file_path}")
        debug_print("ExcelReader - initialized with file:", file_path, "using engine:", self.engine)

    def read_sheets(self):
        """Повертає словник: {sheet_name: DataFrame}."""
        try:
            xls = pd.ExcelFile(self.file_path, engine=self.engine)
            sheets = {}
            for sheet in xls.sheet_names:
                sheets[sheet] = pd.read_excel(xls, sheet_name=sheet)
                debug_print(f"ExcelReader - loaded sheet: {sheet} with shape:", sheets[sheet].shape)
            return sheets
        except Exception as e:
            logging.error(f"Error reading file {self.file_path}: {e}")
            debug_print("ExcelReader - error reading file:", self.file_path, e)
            return {}


class SoilMassParser:
    """Парсить дані ґрунтової маси з Excel та виводить відладочну інформацію."""

    def __init__(self, sheets):
        self.sheets = sheets
        self.data_fields = {
            "Об'ємна маса ґрунту на глибині ґрунту 10 см, г/см3",
            "Об'ємна маса ґрунту на глибині ґрунту 20 см, г/см3"
        }

    def parse(self):
        rows = []
        current_row = {}
        current_date = None
        debug_print("SoilMassParser - початок парсингу.")
        for sheet_name, df in self.sheets.items():
            debug_print("SoilMassParser - обробка аркуша:", sheet_name)
            try:
                for idx, row in df.iterrows():
                    row_text_list = df.iloc[idx].astype(str).dropna().tolist()
                    debug_print(f"SoilMassParser - row {idx} text list:", row_text_list)
                    row_text = ' '.join(row_text_list)
                    new_date_found = None
                    for col_idx, cell in enumerate(row):
                        debug_print(f"SoilMassParser - row {idx} col {col_idx} value:", cell, "type:", type(cell))
                        if pd.notna(cell) and isinstance(cell, str):
                            previous_cell = df.iloc[idx - 1, col_idx] if idx > 0 else None
                            previous_3rd_cell = df.iloc[idx - 3, col_idx] if idx > 2 else None

                            # Парсинг дати
                            if "середня" in cell and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                                date_slice = df.iloc[max(0, idx - 3):idx + 1, 0].dropna()
                                debug_print(f"SoilMassParser - row {idx} date_slice:", date_slice.tolist())
                                if not date_slice.empty:
                                    d = replace_date_format(date_slice.iloc[0], set_day_to_one=True)
                                    if d:
                                        new_date_found = d
                                        debug_print(f"SoilMassParser - row {idx} found new date:", new_date_found)

                            # Парсинг ґрунтової інформації
                            if 'Ґрунт' in cell or 'Грунт' in cell:
                                soil_info = extract_soil_info(cell)
                                if soil_info:
                                    current_row['Ґрунт'] = soil_info
                                    debug_print(f"SoilMassParser - row {idx} extracted ґрунт info from cell:", soil_info)
                                soil_info_list = df.iloc[idx, col_idx + 1:col_idx + 8].dropna().tolist()
                                if soil_info_list:
                                    joined_info = ' '.join(soil_info_list)
                                    current_row['Ґрунт'] = joined_info
                                    debug_print(f"SoilMassParser - row {idx} joined ґрунт info:", joined_info)

                            # Парсинг номеру ділянки
                            if "Ділянка" in cell:
                                plot_info = [str(item) for item in df.iloc[idx, col_idx + 1:col_idx + 10].dropna().tolist()]
                                debug_print(f"SoilMassParser - row {idx} plot_info list:", plot_info)
                                if plot_info:
                                    current_row['Ділянка'] = extract_plot_number(plot_info)
                                else:
                                    current_row['Ділянка'] = extract_plot_number(cell)
                                debug_print(f"SoilMassParser - row {idx} extracted Ділянка:", current_row.get('Ділянка'))

                            # Парсинг показників (наприклад, об'ємна маса ґрунту)
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
                                    debug_print(f"SoilMassParser - row {idx} added data for key '{key}':", data)

                    if new_date_found:
                        if current_date is not None:
                            row_to_add = current_row.copy()
                            row_to_add['Дата'] = current_date
                            if any(field in row_to_add for field in self.data_fields):
                                rows.append(row_to_add)
                                debug_print(f"SoilMassParser - appended row at date change, row:", row_to_add)
                            current_row = {}
                        current_date = new_date_found
                        debug_print(f"SoilMassParser - updated current_date to:", current_date)

                if current_date is not None and current_row:
                    row_to_add = current_row.copy()
                    row_to_add['Дата'] = current_date
                    if any(field in row_to_add for field in self.data_fields):
                        rows.append(row_to_add)
                        debug_print("SoilMassParser - appended final row:", row_to_add)
                current_row = {}
                current_date = None

            except Exception as e:
                logging.error(f"Error reading sheet {sheet_name}: {e}")
                debug_print("SoilMassParser - error in sheet", sheet_name, ":", e)
                continue
        debug_print("SoilMassParser - finished parsing. Total rows:", len(rows))
        return pd.DataFrame(rows)


class MoistureDateParser:
    """
    Парсить дані вологості за рядками DataFrame із детальним виведенням діагностики.
    """

    def __init__(self, df):
        self.df = df

    def parse_sheet_into_df(self):
        df_result = pd.DataFrame()
        data_index = 0
        debug_print("MoistureDateParser - початок парсингу листа.")
        for idx, row in self.df.iterrows():
            row_text_list = self.df.iloc[idx].astype(str).dropna().tolist()
            debug_print(f"MoistureDateParser - row {idx} text list:", row_text_list)
            row_text = ' '.join(row_text_list)
            for col_idx, cell in enumerate(row):
                debug_print(f"MoistureDateParser - row {idx} col {col_idx} value:", cell, "type:", type(cell))
                if pd.notna(cell):
                    previous_cell = self.df.iloc[idx - 1, col_idx] if idx > 0 else None
                    previous_3rd_cell = self.df.iloc[idx - 3, col_idx] if idx > 2 else None

                    # Парсинг дати
                    if "середня" in str(cell) and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                        date_slice = self.df.iloc[max(0, idx - 3):idx + 1, 0].dropna()
                        date_val = date_slice.iloc[-1] if not date_slice.empty else None
                        parsed_date = replace_date_format(date_val, set_day_to_one=False)
                        if parsed_date is not None:
                            data_index = idx
                            df_result.loc[data_index, 'Дата'] = parsed_date
                            debug_print(f"MoistureDateParser - row {idx} found date:", parsed_date)

                    # Парсинг ґрунтової інформації
                    if "Ґрунт" in str(cell) or "Грунт" in str(cell):
                        soil_info_list = self.df.iloc[idx, col_idx + 1:col_idx + 8].dropna().tolist()
                        joined_soil = ' '.join(map(str, soil_info_list))
                        df_result.loc[data_index, 'Ґрунт'] = joined_soil
                        debug_print(f"MoistureDateParser - row {idx} extracted ґрунт info:", joined_soil)

                    # Парсинг ділянки
                    if "Ділянка" in str(cell):
                        plot_info = [str(item) for item in self.df.iloc[idx, col_idx + 1:col_idx + 8].dropna().tolist()]
                        debug_print(f"MoistureDateParser - row {idx} plot_info list:", plot_info)
                        if plot_info:
                            df_result.loc[data_index, 'Ділянка'] = extract_plot_number(plot_info)
                        else:
                            df_result.loc[data_index, 'Ділянка'] = extract_plot_number(cell)
                        debug_print(f"MoistureDateParser - row {idx} extracted Ділянка:", df_result.loc[data_index, 'Ділянка'])

                    # Парсинг середньої вологості
                    if "середня" in str(cell) and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                        average_moisture = self.df.iloc[idx, col_idx + 1:col_idx + 101].dropna().tolist()[:10]
                        avg_columns = [
                            f'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту {i * 10} см, %'
                            for i in range(1, len(average_moisture) + 1)
                        ]
                        df_result.loc[data_index, avg_columns] = average_moisture
                        debug_print(f"MoistureDateParser - row {idx} added середня вологість:", average_moisture)

                    # Парсинг даних по шарах
                    if "по шарах" in str(cell):
                        if "середня" in str(previous_cell):
                            total_moisture_layers = self.df.iloc[idx, col_idx + 1:col_idx + 101].dropna().tolist()[:10]
                            total_moisture_layers = [0 if (isinstance(x, (int, float)) and x < 0) else x for x in total_moisture_layers]
                            total_cols = [
                                f'Запаси загальної вологи по шарах на глибині ґрунту {i * 10} см, мм'
                                for i in range(1, len(total_moisture_layers) + 1)
                            ]
                            df_result.loc[data_index, total_cols] = total_moisture_layers
                            debug_print(f"MoistureDateParser - row {idx} added загальна волога по шарах:", total_moisture_layers)
                        elif "по шарах" in str(previous_cell):
                            productive_moisture_layers = self.df.iloc[idx, col_idx + 1:col_idx + 101].dropna().tolist()[:10]
                            productive_moisture_layers = [0 if (isinstance(x, (int, float)) and x < 0) else x for x in productive_moisture_layers]
                            prod_cols = [
                                f'Запаси продуктивної вологи по шарах на глибині ґрунту {i * 10} см, мм'
                                for i in range(1, len(productive_moisture_layers) + 1)
                            ]
                            df_result.loc[data_index, prod_cols] = productive_moisture_layers
                            debug_print(f"MoistureDateParser - row {idx} added продуктивна волога по шарах:", productive_moisture_layers)

                    # Парсинг наростаючої продуктивної вологості
                    if "наростаючим" in str(cell) and "по шарах" in str(previous_cell):
                        cumulative_moisture = self.df.iloc[idx, col_idx + 1:col_idx + 101].dropna().tolist()[:10]
                        cumulative_moisture = [0 if (isinstance(x, (int, float)) and x < 0) else x for x in cumulative_moisture]
                        cum_cols = [
                            f'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту {i * 10} см, мм'
                            for i in range(1, len(cumulative_moisture) + 1)
                        ]
                        df_result.loc[data_index, cum_cols] = cumulative_moisture
                        debug_print(f"MoistureDateParser - row {idx} added наростаючим волога:", cumulative_moisture)
        debug_print("MoistureDateParser - finished parsing sheet. Resulting rows count:", len(df_result))
        return df_result


class MoistureParser:
    """
    Об'єднує дані вологості з усіх аркушів із додатковою діагностикою.
    """

    def __init__(self, sheets):
        self.sheets = sheets

    def parse(self):
        all_dfs = []
        debug_print("MoistureParser - початок обробки аркушів.")
        for sheet_name, df in self.sheets.items():
            debug_print("MoistureParser - обробка аркуша:", sheet_name)
            parser = MoistureDateParser(df)
            df_moist = parser.parse_sheet_into_df()
            if not df_moist.empty:
                all_dfs.append(df_moist)
                debug_print(f"MoistureParser - аркуш {sheet_name} додано до результату. Форма:", df_moist.shape)
        if all_dfs:
            result_df = pd.concat(all_dfs, ignore_index=True)
            debug_print("MoistureParser - об'єднаний DataFrame має форму:", result_df.shape)
            return result_df
        else:
            debug_print("MoistureParser - не знайдено даних.")
            return pd.DataFrame()


class DataBuilder:
    """Форматує фінальні DataFrame для ґрунтової маси та вологості."""

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

    @staticmethod
    def build_moisture_dataframe(df):
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
        if not df.empty:
            return df.reindex(columns=moisture_columns)
        return df


# --- Основний запуск для тестування одного файлу з детальною діагностикою ---
if __name__ == '__main__':
    # Вкажіть шлях до тестового Excel-файлу
    # test_file_path = r'C:\Users\5302\PycharmProjects\data_drougt\DATA_base_soil_water_cleaned\ТСГ-6_2017\ТСГ-6_2017_Волинська\Любешів\Овес\ТСГ-6_2017_2017_Любешів_03_2017_овес.xls'
    test_file_path = r'C:\Users\5302\PycharmProjects\data_drougt\DATA_base_soil_water_cleaned\ТСГ-6_2016\ТСГ-6_2016_Київська\Яготин\зяб\ТСГ-6_2016_2016_03_2016_Яготин_зяб_під_кукурудзу.xls'

    debug_print("Основний запуск - обробка тестового файлу:", test_file_path)

    try:
        # Читаємо Excel-файл
        reader = ExcelReader(test_file_path)
        sheets = reader.read_sheets()
        debug_print("Основний запуск - завантажені аркуші:", list(sheets.keys()))
        for sheet_name, df in sheets.items():
            debug_print(f"Аркуш '{sheet_name}' має форму:", df.shape)

        # Парсинг даних ґрунтової маси
        debug_print("Основний запуск - початок парсингу ґрунтової маси.")
        soil_parser = SoilMassParser(sheets)
        df_soil = soil_parser.parse()
        debug_print("Основний запуск - результуючий DataFrame для ґрунтової маси:")
        debug_print(df_soil.head(), "\nТип DataFrame:", type(df_soil))

        # Парсинг даних вологості
        debug_print("Основний запуск - початок парсингу вологості.")
        moisture_parser = MoistureParser(sheets)
        df_moist = moisture_parser.parse()
        debug_print("Основний запуск - результуючий DataFrame для вологості:")
        debug_print(df_moist.head(), "\nТип DataFrame:", type(df_moist))

    except Exception as e:
        debug_print("Основний запуск - виникла помилка під час обробки:", e)
        logging.error("Error in main processing: " + str(e))
