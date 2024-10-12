import pandas as pd


def generate_index_map(file_path, default_row_index=14):
    try:
        # Зчитування першого аркуша
        with pd.ExcelFile(file_path, engine='openpyxl') as xls:
            data = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
            column_names = data.columns.tolist()

        # Генерація словника index_map
        index_map = {name: (default_row_index, idx) for idx, name in enumerate(column_names)}
        return index_map
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return {}


# Виклик функції
file_path = './Запаси вологи в ґрунті.xlsx'
index_map = generate_index_map(file_path)
for key, value in index_map.items():
    print(f"'{key}': {value},")
