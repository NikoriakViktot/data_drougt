import os
import logging
import pandas as pd
from datetime import datetime
import re

logging.basicConfig(filename='error_log.txt', level=logging.ERROR)

# ============================================
# Загальні допоміжні функції
# ============================================


def clean_note_value(value: str) -> str | None:
    """
    Очищує вхідний рядок примітки, видаляючи префікс "Примітка" (з або без знаку ":" або "."),
    підкреслення, небажані символи (залишаються лише літери, цифри, крапки та пробіли) і зводить усі пробіли до одного.

    Приклади:
      "Примітка:  _______кінець цвітіння"
          -> "кінець цвітіння"
      "Примітка. Опадомір - на ММ, відстань- 0,2 км. Фаза - н.ф.н. Грунтова засуха."
          -> "Опадомір - на ММ відстань 0,2 км Фаза нфн Грунтова засуха"
      "Примітка: Посадка ___"
          -> "Посадка"
      "Примітка:_"
          -> None
    """
    text = str(value).strip()

    # Якщо рядок після префікса "Примітка" (з або без знаку розділового знаку) складається лише із підкреслень або пробілів,
    # повертаємо None
    if re.match(r'(?i)^примітка[:.\s_]+$', text):
        return None

    # Видаляємо префікс "Примітка" з будь-яким розділовим знаком (":", ".", пробіл)
    text = re.sub(r'(?i)^примітка[:.\s]*', '', text).strip()

    # Замінюємо підкреслення на пробіли
    text = text.replace('_', ' ')

    # Дозволяємо лише літери (українські та латинські), цифри, крапки та пробіли
    text = re.sub(r'[^A-Za-zА-Яа-яЁёІіЇїЄєҐґ0-9.\s]', '', text).strip()

    # Зводимо послідовні пробіли до одного
    text = re.sub(r'\s+', ' ', text).strip()

    # (Опційно) Видаляємо зайві знаки пунктуації з кінця рядка
    text = re.sub(r'[.,:;!?]+$', '', text).strip()

    return text if text else None

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

def replace_date_format(date_text, set_day_to_one=False):
    """
    Зчитує дату із рядка.
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

def extract_plot_number(value):
    """
    Видаляє слово "Ділянка" та символи "№", повертаючи числову частину (якщо є).
    """
    def process(text):
        if not isinstance(text, str):
            text = str(text)
        cleaned = re.sub(r'(?i)\s*Ділянка\s*№?\s*', '', text).strip()
        cleaned = re.sub(r'(?i)^№\s*', '', cleaned).strip()
        return cleaned if re.search(r'\d', cleaned) else None

    if isinstance(value, list):
        for item in value:
            res = process(item)
            if res is not None:
                return res
        return None
    else:
        return process(value)

# ============================================
# ExcelReader – завантаження Excel-файлів
# ============================================

class ExcelReader:
    """Відповідає за завантаження Excel-файлу та отримання даних із аркушів."""
    def __init__(self, file_path):
        self.file_path = file_path
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
            sheets = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
            return sheets
        except Exception as e:
            logging.error(f"Error reading file {self.file_path}: {e}")
            return {}

# ============================================
# Базовий клас для property-стратегій
# ============================================

class BasePropertyStrategy:
    """Базова стратегія для вилучення одного поля з рядка."""
    def __init__(self, property_name: str):
        self.property_name = property_name

    def get_value(self, row, row_idx, df):
        raise NotImplementedError

    def extract(self, row, row_idx, df) -> dict:
        value = self.get_value(row, row_idx, df)
        if value is not None:
            if isinstance(value, dict):
                return value
            return {self.property_name: value}
        return {}

# ============================================
# Конкретні property-стратегії
# ============================================

class DatePropertyStrategy(BasePropertyStrategy):
    def __init__(self, property_name='Дата', set_day_to_one=False):
        super().__init__(property_name)
        self.set_day_to_one = set_day_to_one

    def get_value(self, row, row_idx, df):
        for col_idx, cell in enumerate(row):
            if pd.notna(cell):
                previous_cell = df.iloc[row_idx - 1, col_idx] if row_idx > 0 else None
                previous_3rd_cell = df.iloc[row_idx - 3, col_idx] if row_idx > 2 else None
                if "середня" in str(cell) and ("4" in str(previous_cell) or "2" in str(previous_3rd_cell)):
                    date_slice = df.iloc[max(0, row_idx - 3):row_idx + 1, 0].dropna()
                    date_val = date_slice.iloc[-1] if not date_slice.empty else None
                    return replace_date_format(date_val, set_day_to_one=self.set_day_to_one)
        return None

class PlotPropertyStrategy(BasePropertyStrategy):
    def __init__(self, property_name='Ділянка'):
        super().__init__(property_name)
        self.cached_plot = None

    def reset_cache(self):
        """Скидає кешований номер ділянки, щоб при новому файлі шукати заново."""
        self.cached_plot = None

    def get_value(self, row, row_idx, df):
        if self.cached_plot is not None:
            return self.cached_plot
        for col_idx, cell in enumerate(row):
            if pd.notna(cell) and "ділянка" in str(cell).lower():
                plot_info = df.iloc[row_idx, col_idx+1:col_idx+10].dropna().astype(str).tolist()
                if plot_info:
                    self.cached_plot = extract_plot_number(plot_info)
                    return self.cached_plot
                else:
                    self.cached_plot = extract_plot_number(cell)
                    return self.cached_plot
        return None




# ------------------------------------------------------------------
# клас SoilPropertyStrategy з очищенням даних ґрунту
# ------------------------------------------------------------------

class SoilPropertyStrategy(BasePropertyStrategy):
    def __init__(self, property_name='Ґрунт'):
        super().__init__(property_name)
        self.cached_soil = None

    def reset_cache(self):
        """Скидає кешований тип ґрунту, щоб при новому файлі шукати заново."""
        self.cached_soil = None

    def get_value(self, row, row_idx, df):
        if self.cached_soil is not None:
            return self.cached_soil
        for col_idx, cell in enumerate(row):
            if pd.notna(cell) and ("ґрунт" in str(cell).lower() or "грунт" in str(cell).lower()):
                soil_info_list = df.iloc[row_idx, col_idx + 1:col_idx + 8].dropna().astype(str).tolist()
                if soil_info_list:
                    combined = ' '.join(soil_info_list)
                    self.cached_soil = clean_soil_string(combined)
                    return self.cached_soil
        return None


class AverageMoisturePropertyStrategy(BasePropertyStrategy):
    def __init__(self, property_name_prefix='Вологість від абсолютно сухого ґрунту (середня)'):
        super().__init__(property_name_prefix)
        self.property_name_prefix = property_name_prefix

    def get_value(self, row, row_idx, df):
        for col_idx, cell in enumerate(row):
            if pd.notna(cell) and "середня" in str(cell):
                previous_cell = df.iloc[row_idx-1, col_idx] if row_idx > 0 else None
                previous_3rd_cell = df.iloc[row_idx-3, col_idx] if row_idx > 2 else None
                if "4" in str(previous_cell) or "2" in str(previous_3rd_cell):
                    avg_values = df.iloc[row_idx, col_idx+1:col_idx+101].dropna().tolist()[:10]
                    result = {}
                    for i, value in enumerate(avg_values, start=1):
                        col_name = f'{self.property_name_prefix} на глибині ґрунту {i*10} см, %'
                        result[col_name] = value
                    return result
        return {}

class TotalMoisturePropertyStrategy(BasePropertyStrategy):
    def __init__(self, property_name_prefix='Запаси загальної вологи по шарах'):
        super().__init__(property_name_prefix)
        self.property_name_prefix = property_name_prefix

    def get_value(self, row, row_idx, df):
        for col_idx, cell in enumerate(row):
            if pd.notna(cell) and "по шарах" in str(cell).lower():
                previous_cell = df.iloc[row_idx-1, col_idx] if row_idx > 0 else None
                if previous_cell is not None and "середня" in str(previous_cell).lower():
                    values = df.iloc[row_idx, col_idx+1:col_idx+101].dropna().tolist()[:10]
                    values = [0 if (isinstance(x, (int, float)) and x < 0) else x for x in values]
                    result = {}
                    for i, value in enumerate(values, start=1):
                        col_name = f'{self.property_name_prefix} на глибині ґрунту {i*10} см, мм'
                        result[col_name] = value
                    return result
        return {}

class ProductiveMoisturePropertyStrategy(BasePropertyStrategy):
    def __init__(self, property_name_prefix='Запаси продуктивної вологи по шарах'):
        super().__init__(property_name_prefix)
        self.property_name_prefix = property_name_prefix

    def get_value(self, row, row_idx, df):
        for col_idx, cell in enumerate(row):
            if pd.notna(cell) and "по шарах" in str(cell).lower():
                previous_cell = df.iloc[row_idx-1, col_idx] if row_idx > 0 else None
                if previous_cell is not None and "по шарах" in str(previous_cell).lower():
                    values = df.iloc[row_idx, col_idx+1:col_idx+101].dropna().tolist()[:10]
                    values = [0 if (isinstance(x, (int, float)) and x < 0) else x for x in values]
                    result = {}
                    for i, value in enumerate(values, start=1):
                        col_name = f'{self.property_name_prefix} на глибині ґрунту {i*10} см, мм'
                        result[col_name] = value
                    return result
        return {}

class CumulativeMoisturePropertyStrategy(BasePropertyStrategy):
    def __init__(self, property_name_prefix='Запаси продуктивної вологи наростаючим підсумком'):
        super().__init__(property_name_prefix)
        self.property_name_prefix = property_name_prefix

    def get_value(self, row, row_idx, df):
        for col_idx, cell in enumerate(row):
            if pd.notna(cell) and "наростаючим" in str(cell).lower():
                previous_cell = df.iloc[row_idx-1, col_idx] if row_idx > 0 else None
                if previous_cell is not None and "по шарах" in str(previous_cell).lower():
                    values = df.iloc[row_idx, col_idx+1:col_idx+101].dropna().tolist()[:10]
                    values = [0 if (isinstance(x, (int, float)) and x < 0) else x for x in values]
                    result = {}
                    for i, value in enumerate(values, start=1):
                        col_name = f'{self.property_name_prefix} на глибині ґрунту {i*10} см, мм'
                        result[col_name] = value
                    return result
        return {}


class NotePropertyStrategy(BasePropertyStrategy):
    def __init__(self, property_name='Примітка'):
        super().__init__(property_name)

    def get_value(self, row, row_idx, df):
        """
        Шукає в кожному рядку згадку про 'Примітка' (у будь-якому стовпці).
        Якщо знайдено, очищує та повертає текст.
        """
        for col_idx, cell in enumerate(row):
            if pd.notna(cell):
                text_lower = str(cell).lower()
                if 'примітка' in text_lower:
                    # Викликаємо вашу функцію clean_note_value
                    cleaned_note = clean_note_value(str(cell))
                    if cleaned_note:
                        return cleaned_note
        return None


# ============================================
# Data Builder, який координує виклики property-стратегій
# ============================================

class MoistureDataBuilder:
    """
    Координує виклики property-стратегій для побудови даних з одного рядка.
    """
    def __init__(self, strategies):
        # strategies – список об'єктів, що наслідують BasePropertyStrategy
        self.strategies = strategies

    def build_row(self, row, row_idx, df) -> dict:
        data = {}
        for strategy in self.strategies:
            extracted = strategy.extract(row, row_idx, df)
            if extracted:
                data.update(extracted)
        return data

# ============================================
# Блочний парсер, що групує дані за днями
# ============================================

class DayBlockParser:
    """
    Парсер, який проходить по рядках DataFrame та групує їх за днями.
    Якщо стратегія DatePropertyStrategy повертає дату, і ця дата відрізняється від поточної,
    починається новий «блок» даних (новий день).
    """
    def __init__(self, df, data_builder):
        self.df = df
        self.data_builder = data_builder

    def parse_sheet(self):
        all_days_data = []
        current_day_data = {}
        current_date = None

        for idx, row in self.df.iterrows():
            row_data = self.data_builder.build_row(row, idx, self.df)

            # Якщо в row_data знайшлась дата
            if 'Дата' in row_data:
                new_date = row_data['Дата']

                # Якщо була поточна дата, зберігаємо блок
                if current_date is not None:
                    all_days_data.append(current_day_data)

                # Починаємо новий блок
                current_day_data = {}
                current_date = new_date

            # Якщо є поточна дата, додаємо дані
            if current_date is not None:
                for key, value in row_data.items():
                    # Перезаписуємо лише якщо ключ ще не встановлено
                    if key not in current_day_data or current_day_data[key] in [None, '']:
                        current_day_data[key] = value

        # Після закінчення циклу, якщо є блок із датою, додаємо
        if current_date is not None:
            all_days_data.append(current_day_data)

        return pd.DataFrame(all_days_data)


# ============================================
# Парсер вологості, що об'єднує дані з усіх аркушів
# ============================================

class MoistureParser:
    """
    Приймає словник {sheet_name: DataFrame},
    використовує DayBlockParser для групування рядків за днями,
    і об'єднує результати в один DataFrame.
    """
    def __init__(self, sheets, data_builder):
        self.sheets = sheets
        self.data_builder = data_builder

    def parse(self):
        all_dfs = []
        for sheet_name, df in self.sheets.items():
            parser = DayBlockParser(df, self.data_builder)
            df_moist = parser.parse_sheet()
            if not df_moist.empty:
                all_dfs.append(df_moist)
        if all_dfs:
            return pd.concat(all_dfs, ignore_index=True)
        return pd.DataFrame()

# ============================================
# Фінальне форматування DataFrame
# ============================================

class DataBuilder:
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
            'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 100 см, мм',
            'Примітка'
       ]
        if not df.empty:
            return df.reindex(columns=moisture_columns)
        return df

# ============================================
# DataProcessor – оркестрація обробки файлів
# ============================================

class DataProcessor:
    """
    Оркеструє процес обробки файлів та побудови фінальної структури даних.
    """
    def __init__(self, base_folder, years, data_builder):
        self.base_folder = base_folder
        self.years = years
        self.data_builder = data_builder
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
                                    # Перед обробкою нового файлу скидаємо кеш
                                    for strategy in self.data_builder.strategies:
                                        if hasattr(strategy, 'reset_cache'):
                                            strategy.reset_cache()

                                    # Далі зчитуємо sheets
                                    reader = ExcelReader(file_path)
                                    sheets = reader.read_sheets()

                                    # Парсимо дані
                                    moisture_parser = MoistureParser(sheets, self.data_builder)
                                    df_moist = moisture_parser.parse()

                                    if not df_moist.empty:
                                        df_moist['file_path'] = file_path
                                        df_moist['Рік'] = year
                                        df_moist['Область'] = folder_obl[11:]
                                        df_moist['Метеостанція'] = folder_st
                                        df_moist['Культура'] = folder_crop.lower()
                                        self.moisture_data = pd.concat([self.moisture_data, df_moist],
                                                                       ignore_index=True)

                                except Exception as e:
                                    logging.error(f"Error processing file {file_path}: {e}")

        self.moisture_data.dropna(axis=1, how='all', inplace=True)

    def save(self, moisture_output_file):
        moisture_df = DataBuilder.build_moisture_dataframe(self.moisture_data)
        moisture_df.to_excel(moisture_output_file, index=False)
        print(f"Processing completed.\nMoisture data -> {moisture_output_file}")

# ============================================
# Основний запуск
# ============================================

if __name__ == '__main__':
    base_folder = r"C:\Users\5302\PycharmProjects\data_drougt\DATA_base_soil_water_cleaned"
    years = ['2016', '2017', '2018', '2019', '2020', '2021']
    property_strategies = [
        DatePropertyStrategy(set_day_to_one=False),
        PlotPropertyStrategy(),
        SoilPropertyStrategy(),
        AverageMoisturePropertyStrategy(),
        TotalMoisturePropertyStrategy(),
        ProductiveMoisturePropertyStrategy(),
        CumulativeMoisturePropertyStrategy(),
        NotePropertyStrategy()
    ]
    data_builder = MoistureDataBuilder(property_strategies)
    processor = DataProcessor(base_folder, years, data_builder)
    processor.process()
    processor.save('DB_Moisture_ver_1_1.xlsx')
