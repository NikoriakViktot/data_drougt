import pandas as pd
import os

# === 1. Введення шляхів до файлів ===
file1_path = r"C:\Users\5302\PycharmProjects\data_drougt\DB_Moisture_ver_1_1.xlsx"  # Змініть на свій шлях до першого файлу
file2_path = r"C:\Users\5302\PycharmProjects\data_drougt\DB_Soil_mass_ver_1_1.xlsx"  # Змініть на свій шлях до другого файлу
output_file = "merged_DB_soil_moisture_ver_1.xlsx"  # Ім'я вихідного файлу

# === 2. Функція для завантаження файлів ===
def load_file(file_path):
    """ Завантажує файл у DataFrame, підтримує CSV та Excel. """
    ext = os.path.splitext(file_path)[-1].lower()
    if ext == ".csv":
        return pd.read_csv(file_path, encoding='utf-8')
    elif ext in [".xls", ".xlsx"]:
        return pd.read_excel(file_path)
    else:
        raise ValueError(f"Непідтримуваний формат файлу: {file_path}")

# Завантаження обох файлів
df1 = load_file(file1_path)
df2 = load_file(file2_path)

# === 3. Визначення спільних ключових колонок ===
common_columns = ['file_path', 'Рік', 'Область', 'Метеостанція', 'Ділянка', 'Культура', 'Ґрунт', 'Дата']

# === 4. Об'єднання файлів ===
merged_df = pd.merge(df1, df2, on=common_columns, how='outer')

# Перетворення дати в коректний формат
merged_df['Дата'] = pd.to_datetime(merged_df['Дата'], errors='coerce')

# === 5. Збереження результату ===
output_ext = os.path.splitext(output_file)[-1].lower()

if output_ext == ".csv":
    merged_df.to_csv(output_file, index=False, encoding='utf-8')
elif output_ext in [".xls", ".xlsx"]:
    merged_df.to_excel(output_file, index=False)
else:
    raise ValueError(f"Непідтримуваний формат файлу для збереження: {output_file}")

print(f"Файл успішно збережено: {output_file}")
