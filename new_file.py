import pandas as pd
import re
from datetime import datetime

# Function to replace various date formats with a standard datetime format
def replace_date_format(date_text):
    # Якщо вже є datetime об'єкт, просто повертаємо його
    if isinstance(date_text, datetime):
        return date_text

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
        match = re.search(pattern, str(date_text))  # перевірка на строки
        if match:
            day, month, year = match.groups()
            if len(year) == 2:
                year = '20' + year
            try:
                # Convert string to datetime object
                return datetime.strptime(f"{year}-{month}-{day}", "%Y-%m-%d")
            except ValueError:
                continue  # If conversion fails, try next pattern
    return date_text  # If no pattern matches, return None

# Load the DataFrame from an Excel file
df = pd.read_excel('DB_Soil_water_2016_2021_raw_3.xlsx')

# Apply the date correction to the 'Дата' column
df['Дата'] = df['Дата'].apply(lambda x: replace_date_format(str(x)) if pd.notna(x) else None)

# Save the cleaned DataFrame back to an Excel file
df.to_excel('DB_Soil_water_2016_2021_cleaned_date.xlsx', index=False)
