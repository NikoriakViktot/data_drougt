import pandas as pd
import re
from datetime import datetime

def replace_date_format(date_text):
    date_patterns = [
        r'(\d{2})/(\d{2})/(\d{4})',     # 18/07/2016
        r'(\d{2})\.(\d{2})\.(\d{4})\.', # 18.07.2016.
        r'(\d{2})\.(\d{2})\.(\d{4})',   # 08.07.2015
        r'(\d{2})\.(\d{2})\.(\d{2})\sр\.',  # 28.07.16 р.
        r'(\d{2})\.(\d{2})\.(\d{2})\sр',    # 27.04.16 р
        r'(\d{2})\.(\d{2})\.(\d{2})'    # 27.04.16
    ]
    for pattern in date_patterns:
        match = re.search(pattern, date_text)
        if match:
            day, month, year = match.groups()
            if len(year) == 2:
                year = '20' + year
            return f"{day}/{month}/{year}"  # Returning in YYYY-MM-DD format
    return date_text  # Return original if no pattern matches

def get_date_from_row(data, primary_col, primary_row_range, secondary_col, secondary_row_range):
    # Check in the primary column
    for row_index in range(primary_row_range[0], primary_row_range[1] + 1):
        date = data.loc[row_index, primary_col]
        if pd.notna(date):
            # Check if the date is already a datetime object
            if not isinstance(date, datetime):
                date = replace_date_format(str(date))
            return date

    # Check in the secondary column if no valid date was found in the primary column
    for row_index in range(secondary_row_range[0], secondary_row_range[1] + 1):
        date = data.loc[row_index, secondary_col]
        if pd.notna(date):
            # Check if the date is already a datetime object
            if not isinstance(date, datetime):
                date = replace_date_format(str(date))
            return date

    return None  # Return None if no valid dates were found

# Example DataFrame to use for testing
data = pd.DataFrame({
    'A': [None, '27.04.16', '2021-01-01', '15/12/2020'],
    'B': ['2020-05-20', None, '08.02.16', '28.12.2020.']
})

print(get_date_from_row(data, 'A', (0, 3), 'B', (0, 3)))
import pandas as pd
import re
from datetime import datetime

def replace_date_format(date_text):
    date_patterns = [
        r'(\d{2})/(\d{2})/(\d{4})',     # 18/07/2016
        r'(\d{2})\.(\d{2})\.(\d{4})\.', # 18.07.2016.
        r'(\d{2})\.(\d{2})\.(\d{4})',   # 08.07.2015
        r'(\d{2})\.(\d{2})\.(\d{2})\sр\.',  # 28.07.16 р.
        r'(\d{2})\.(\d{2})\.(\d{2})\sр',    # 27.04.16 р
        r'(\d{2})\.(\d{2})\.(\d{2})'    # 27.04.16
    ]
    for pattern in date_patterns:
        match = re.search(pattern, date_text)
        if match:
            day, month, year = match.groups()
            if len(year) == 2:
                year = '20' + year
            return f"{year}-{month}-{day}"  # Returning in YYYY-MM-DD format
    return date_text  # Return original if no pattern matches

def get_dates_from_rows(data, primary_col, primary_row_range, secondary_col, secondary_row_range):
    valid_dates = []
    # Check in the primary column
    for row_index in range(primary_row_range[0], primary_row_range[1] + 1):
        date = data.loc[row_index, primary_col]
        if pd.notna(date):
            if not isinstance(date, datetime):
                date = replace_date_format(str(date))
            valid_dates.append(date)

    # Check in the secondary column if no valid date was found in the primary column
    for row_index in range(secondary_row_range[0], secondary_row_range[1] + 1):
        date = data.loc[row_index, secondary_col]
        if pd.notna(date):
            if not isinstance(date, datetime):
                date = replace_date_format(str(date))
            valid_dates.append(date)

    return valid_dates  # Return list of valid dates

# Example DataFrame to use for testing
data = pd.DataFrame({
    'A': [None, '27.04.16', '2021-01-01', '15/12/2020'],
    'B': ['2020-05-20', None, '08.02.16', '28.12.2020.']
})

print(get_dates_from_rows(data, 'A', (0, 3), 'B', (0, 3)))
