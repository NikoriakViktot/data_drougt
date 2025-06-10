import pandas as pd

# Завантаження вашого основного датасету
df = pd.read_excel('DB_Moisture_ver_1_1.xlsx')
print(df.columns.tolist())


# 2. Перекодування полів, де пошкоджені символи
def fix_encoding(x):
    if isinstance(x, str):
        try:
            return x.encode('latin1').decode('utf-8')
        except UnicodeEncodeError:
            return x
    return x

for col in ['Метеостанція', 'Культура' , 'Ґрунт']:
    df[col] = df[col].apply(fix_encoding)



# Завантаження таблиці відповідностей обласних назв
region_mapping = {
    'Вінницька': "Vinnytska oblast",
    'Волинська': "Volynska oblast",
    'Дніпропетровська': "Dnipropetrovska oblast",
    'Донецька': "Donetska oblast",
    'Житомирська': "Zhytomyrska oblast",
    'Закарпатська': "Zakarpatska oblast",
    'Запорізька': "Zaporizka oblast",
    'Івано-Франківська': "Ivano-Frankivska oblast",
    'Київська': "Kyivska oblast",
    'Кіровоградська': "Kirovohradska oblast",
    'Луганська': "Luhanska oblast",
    'Львівська': "Lvivska oblast",
    'Миколаївська': "Mykolaivska oblast",
    'Одеська': "Odeska oblast",
    'Полтавська': "Poltavska oblast",
    'Рівненська': "Rivnenska oblast",
    'Сумська': "Sumska oblast",
    'Тернопільська': "Ternopilska oblast",
    'Харківська': "Kharkivska oblast",
    'Херсонська': "Khersonska oblast",
    'Хмельницька': "Khmelnytska oblast",
    'Черкаська': "Cherkaska oblast",
    'Чернівецька': "Chernivetska oblast",
    'Чернігівська': "Chernihivska oblast",
    'Автономна Республіка Крим': "Autonomous Republic of Crimea"  # якщо потрібно
}

# Список колонок вологості
moisture_columns = [
    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 10 см, %',
    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 20 см, %',
    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 30 см, %',
    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 40 см, %',
    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 50 см, %',
    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 60 см, %',
    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 70 см, %',
    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 80 см, %',
    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 90 см, %',
    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 100 см, %'
]

# Вибір потрібних колонок
selected_columns = ['Рік', 'Область', 'Метеостанція', 'Ділянка', 'Культура', 'Ґрунт', 'Дата'] + moisture_columns + ['Примітка']
df_selected = df[selected_columns]

# Обчислення середньої вологості
df_selected['Average_soil_moisture'] = df_selected[moisture_columns].mean(axis=1, skipna=True)

# Визначення наявності посухи
def detect_drought(note):
    if isinstance(note, str):
        note_lower = note.lower()
        if 'посуха' in note_lower or 'засуха' in note_lower:
            return 1
    return 0

df_selected['Drought'] = df_selected['Примітка'].apply(detect_drought)

# Перейменування назв колонок на англійські
df_selected = df_selected.rename(columns={
    'Рік': 'Year',
    'Область': 'Region',
    'Метеостанція': 'Meteorological_station',
    'Ділянка': 'Plot',
    'Культура': 'Crop',
    'Ґрунт': 'Soil',
    'Дата': 'Date'
})

# Переклад назв областей
df_selected['Region'] = df_selected['Region'].map(region_mapping)

# Формування фінального датасету
df_final = df_selected[['Year', 'Region', 'Meteorological_station', 'Plot', 'Crop', 'Soil', 'Date', 'Average_soil_moisture', 'Drought']]

# Збереження результату
df_final.to_csv('dataset_for_DCS.csv', index=False)

print("✅ Done! File 'dataset_for_DCS.csv' has been created with English column and region names.")
