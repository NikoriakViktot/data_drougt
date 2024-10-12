
import pandas as pd
import re

# Завантажте ваш файл Excel
xls_data = pd.read_excel('your_file.xlsx')

# Функція для пошуку
def find_soil_info(cell):
    if pd.notna(cell):  # Перевірка на наявність значення в клітинці
        cell_str = str(cell)
        # Регулярний вираз для знаходження "Ґрунт" або "Грунт" з будь-яким текстом після нього
        pattern = r'\b(Ґрунт|Грунт)\b.*'
        match = re.search(pattern, cell_str, re.IGNORECASE)
        if match:
            return match.group(0)  # Повертає знайдений текст
    return None

# Перебір всіх клітинок у DataFrame
results = []
for idx, row in xls_data.iterrows():
    for col_idx, cell in enumerate(row):
        soil_info = find_soil_info(cell)
        if soil_info:
            # Можна зберігати результат або використовувати його за потребою
            results.append((idx, col_idx, soil_info))

# Перевірка результатів
for result in results:
    print(f"Index: {result[0]}, Column: {result[1]}, Soil Info: {result[2]}")
