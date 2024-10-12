import pandas as pd
import os
import logging

from new_reder_soil_data import SoilData

# logging.basicConfig(filename='process_directory.log', level=logging.INFO,
#                     format='%(asctime)s:%(levelname)s:%(message)s')

logging.basicConfig(filename='dataframebuilder.log', level=logging.INFO,
                    format='%(asctime)s:%(levelname)s:%(message)s')

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

class DataFrameBuilder:

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


    def __init__(self, soil_data_file):
        self.data_frame = pd.DataFrame(columns=self.column_new_data)
        self.soil_data = None


    def set_soil_data(self, soil_data_file):
        self.soil_data = SoilData(soil_data_file)
        self.add_data()

    def reset_data_frame(self):
        self.data_frame = pd.DataFrame(columns=self.column_new_data)

    def add_data_1_step(self):
        try:
            new_row = {
                'file_path': self.soil_data.name_file,
                'Рік': self.soil_data.year,
                'Область': self.soil_data.region,
                'Метеостанція': self.soil_data.weather_station,
                'Культура': self.soil_data.culture,
                'Попередня культура': self.soil_data.previous_culture,
                'Тип ґрунту': self.soil_data.soil_type,
                'Ділянка': self.soil_data.plot,
            }
            new_row_df = pd.DataFrame([new_row])
            row_1 = self.soil_data.step_1
            self.data_frame = pd.concat([self.data_frame, new_row_df, row_1], ignore_index=True)
            logging.info("Data added successfully to DataFrame.")
        except Exception as e:
            logging.error(f"Failed to add data to DataFrame: {e}")
        print("Data added successfully. Current DataFrame:-----> step_1")

    def add_data_1_date(self):
        try:
            new_row = {
                'Рік': self.soil_data.year,
                'Область': self.soil_data.region,
                'Метеостанція': self.soil_data.weather_station,
                'Дата': self.soil_data.date_1,
                'Культура': self.soil_data.culture,
                'Попередня культура': self.soil_data.previous_culture,
                'Тип ґрунту': self.soil_data.soil_type,
                'Ділянка': self.soil_data.plot,
            }
            new_row_df = pd.DataFrame([new_row])  # Створюємо DataFrame з одного рядка
            row_2 = self.soil_data.step_2 if not self.soil_data.step_2.empty else pd.DataFrame()
            self.data_frame = pd.concat([self.data_frame, new_row_df, row_2],
                                        ignore_index=True)  # Додаємо рядок до основного DataFrame
            logging.info("Data added successfully to DataFrame for the second date.")
        except Exception as e:
            logging.error(f"Failed to add data to DataFrame for the second date: {e}")

        print("Data added successfully for the second date. Current DataFrame: ------> 1")

    def add_data_2_date(self):
        try:
            new_row = {
                'Рік': self.soil_data.year,
                'Область': self.soil_data.region,
                'Метеостанція': self.soil_data.weather_station,
                'Дата': self.soil_data.date_2,
                'Культура': self.soil_data.culture,
                'Попередня культура': self.soil_data.previous_culture,
                'Тип ґрунту': self.soil_data.soil_type,
                'Ділянка': self.soil_data.plot,
            }
            new_row_df = pd.DataFrame([new_row])  # Створюємо DataFrame з одного рядка
            row_2 = self.soil_data.step_3 if not self.soil_data.step_3.empty else pd.DataFrame()
            self.data_frame = pd.concat([self.data_frame, new_row_df, row_2],
                                        ignore_index=True)  # Додаємо рядок до основного DataFrame
            logging.info("Data added successfully to DataFrame for the second date.")
        except Exception as e:
            logging.error(f"Failed to add data to DataFrame for the second date: {e}")

        print("Data added successfully for the second date. Current DataFrame: ------> 2")



    def add_data_3_date(self):
        try:
            new_row = {
                'Рік': self.soil_data.year,
                'Область': self.soil_data.region,
                'Метеостанція': self.soil_data.weather_station,
                'Дата': self.soil_data.date_3,
                'Культура': self.soil_data.culture,
                'Попередня культура': self.soil_data.previous_culture,
                'Тип ґрунту': self.soil_data.soil_type,
                'Ділянка': self.soil_data.plot,
            }
            new_row_df = pd.DataFrame([new_row])  # Створюємо DataFrame з одного рядка
            row_2 = self.soil_data.step_4 if not self.soil_data.step_4.empty else pd.DataFrame()
            self.data_frame = pd.concat([self.data_frame, new_row_df, row_2],
                                        ignore_index=True)  # Додаємо рядок до основного DataFrame
            logging.info("Data added successfully to DataFrame for the second date.")
        except Exception as e:
            logging.error(f"Failed to add data to DataFrame for the second date: {e}")

        print("Data added successfully for the second date. Current DataFrame: ------> 3")


    def add_data(self):
        self.add_data_1_step()
        self.add_data_1_date()
        self.add_data_2_date()
        self.add_data_3_date()

    def get_result(self):
        return self.data_frame

class ExcelDataDirector:
    def __init__(self, builder):
        self.builder = builder

    def construct(self, file_path):
        self.builder.reset_data_frame()
        self.builder.set_soil_data(file_path)
        return self.builder.get_result()


def process_directory(directory, output_file='combined_data.xlsx'):
    master_df = pd.DataFrame()
    builder = DataFrameBuilder(None)
    director = ExcelDataDirector(builder)

    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(('.xls', '.xlsx')):
                file_path = os.path.join(root, file)
                logging.info(f"Attempting to open file: {file_path}")
                try:
                    result_df = director.construct(file_path)
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

process_directory(r"C:\Users\user\PycharmProjects\data_drought\DATA_base_soil_water", 'DB_soil_data_16.xlsx')

# column_new_data = ['file_path',
#                    'Рік',
#                    'Область',
#                    'Метеостанція',
#                    # 'Місяць',
#                    'Дата',
#                    'Культура',
#                    'Попередня культура',
#                    'Тип ґрунту',
#                    'Ділянка',
#                    "Об'ємна маса ґрунту на глибині ґрунту 10 см, г/ см³",
#                    'Об\'ємна маса ґрунту на глибині ґрунту 20 см, г/ см³',
#                    'Об\'ємна маса ґрунту на глибині ґрунту 30 см, г/ см³',
#                    'Об\'ємна маса ґрунту на глибині ґрунту 40 см, г/ см³',
#                    'Об\'ємна маса ґрунту на глибині ґрунту 50 см, г/ см³',
#                    'Об\'ємна маса ґрунту на глибині ґрунту 60 см, г/ см³',
#                    'Об\'ємна маса ґрунту на глибині ґрунту 70 см, г/ см³',
#                    'Об\'ємна маса ґрунту на глибині ґрунту 80 см, г/ см³',
#                    'Об\'ємна маса ґрунту на глибині ґрунту 90 см, г/ см³',
#                    'Об\'ємна маса ґрунту на глибині ґрунту 100 см, г/ см³',
#                    'Запаси непродуктивної вологи на глибині ґрунту 10 см, мм',
#                    'Запаси непродуктивної вологи на глибині ґрунту 20 см, мм',
#                    'Запаси непродуктивної вологи на глибині ґрунту 30 см, мм',
#                    'Запаси непродуктивної вологи на глибині ґрунту 40 см, мм',
#                    'Запаси непродуктивної вологи на глибині ґрунту 50 см, мм',
#                    'Запаси непродуктивної вологи на глибині ґрунту 60 см, мм',
#                    'Запаси непродуктивної вологи на глибині ґрунту 70 см, мм',
#                    'Запаси непродуктивної вологи на глибині ґрунту 80 см, мм',
#                    'Запаси непродуктивної вологи на глибині ґрунту 90 см, мм',
#                    'Запаси непродуктивної вологи на глибині ґрунту 100 см, мм',
#                    'Запаси продуктивної вологи при НВ на глибині ґрунту 10 см, мм',
#                    'Запаси продуктивної вологи при НВ на глибині ґрунту 20 см, мм',
#                    'Запаси продуктивної вологи при НВ на глибині ґрунту 30 см, мм',
#                    'Запаси продуктивної вологи при НВ на глибині ґрунту 40 см, мм',
#                    'Запаси продуктивної вологи при НВ на глибині ґрунту 50 см, мм',
#                    'Запаси продуктивної вологи при НВ на глибині ґрунту 60 см, мм',
#                    'Запаси продуктивної вологи при НВ на глибині ґрунту 70 см, мм',
#                    'Запаси продуктивної вологи при НВ на глибині ґрунту 80 см, мм',
#                    'Запаси продуктивної вологи при НВ на глибині ґрунту 90 см, мм',
#                    'Запаси продуктивної вологи при НВ на глибині ґрунту 100 см, мм',
#                    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 10 см, %',
#                    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 20 см, %',
#                    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 30 см, %',
#                    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 40 см, %',
#                    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 50 см, %',
#                    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 60 см, %',
#                    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 70 см, %',
#                    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 80 см, %',
#                    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 90 см, %',
#                    'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 100 см, %',
#                    'Запаси загальної вологи по шарах на глибині ґрунту 10 см, мм',
#                    'Запаси загальної вологи по шарах на глибині ґрунту 20 см, мм',
#                    'Запаси загальної вологи по шарах на глибині ґрунту 30 см, мм',
#                    'Запаси загальної вологи по шарах на глибині ґрунту 40 см, мм',
#                    'Запаси загальної вологи по шарах на глибині ґрунту 50 см, мм',
#                    'Запаси загальної вологи по шарах на глибині ґрунту 60 см, мм',
#                    'Запаси загальної вологи по шарах на глибині ґрунту 70 см, мм',
#                    'Запаси загальної вологи по шарах на глибині ґрунту 80 см, мм',
#                    'Запаси загальної вологи по шарах на глибині ґрунту 90 см, мм',
#                    'Запаси загальної вологи по шарах на глибині ґрунту 100 см, мм',
#                    'Запаси продуктивної вологи по шарах на глибині ґрунту 10 см, мм',
#                    'Запаси продуктивної вологи по шарах на глибині ґрунту 20 см, мм',
#                    'Запаси продуктивної вологи по шарах на глибині ґрунту 30 см, мм',
#                    'Запаси продуктивної вологи по шарах на глибині ґрунту 40 см, мм',
#                    'Запаси продуктивної вологи по шарах на глибині ґрунту 50 см, мм',
#                    'Запаси продуктивної вологи по шарах на глибині ґрунту 60 см, мм',
#                    'Запаси продуктивної вологи по шарах на глибині ґрунту 70 см, мм',
#                    'Запаси продуктивної вологи по шарах на глибині ґрунту 80 см, мм',
#                    'Запаси продуктивної вологи по шарах на глибині ґрунту 90 см, мм',
#                    'Запаси продуктивної вологи по шарах на глибині ґрунту 100 см, мм',
#                    'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 10 см, мм',
#                    'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 20 см, мм',
#                    'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 30 см, мм',
#                    'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 40 см, мм',
#                    'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 50 см, мм',
#                    'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 60 см, мм',
#                    'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 70 см, мм',
#                    'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 80 см, мм',
#                    'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 90 см, мм',
#                    'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 100 см, мм']
#                    # 'Серередня температура повітря за період спостереження, С',
#                    # 'Сума опадів за період спостереження, мм',
#                    # 'Примітка']

# "Об'ємна маса ґрунту на глибині ґрунту 10 см, г/ см³": self.soil_data.soil_density_10cm,
# 'Об\'ємна маса ґрунту на глибині ґрунту 20 см, г/ см³': self.soil_data.soil_density_20cm,
# 'Об\'ємна маса ґрунту на глибині ґрунту 30 см, г/ см³': self.soil_data.soil_density_30cm,
# 'Об\'ємна маса ґрунту на глибині ґрунту 40 см, г/ см³': self.soil_data.soil_density_40cm,
# 'Об\'ємна маса ґрунту на глибині ґрунту 50 см, г/ см³': self.soil_data.soil_density_50cm,
# 'Об\'ємна маса ґрунту на глибині ґрунту 60 см, г/ см³': self.soil_data.soil_density_60cm,
# 'Об\'ємна маса ґрунту на глибині ґрунту 70 см, г/ см³': self.soil_data.soil_density_70cm,
# 'Об\'ємна маса ґрунту на глибині ґрунту 80 см, г/ см³': self.soil_data.soil_density_80cm,
# 'Об\'ємна маса ґрунту на глибині ґрунту 90 см, г/ см³': self.soil_data.soil_density_90cm,
# 'Об\'ємна маса ґрунту на глибині ґрунту 100 см, г/ см³': self.soil_data.soil_density_100cm,
# 'Запаси непродуктивної вологи на глибині ґрунту 10 см, мм': self.soil_data.unproductive_moisture_10cm,
# 'Запаси непродуктивної вологи на глибині ґрунту 20 см, мм': self.soil_data.unproductive_moisture_20cm,
# 'Запаси непродуктивної вологи на глибині ґрунту 30 см, мм': self.soil_data.unproductive_moisture_30cm,
# 'Запаси непродуктивної вологи на глибині ґрунту 40 см, мм': self.soil_data.unproductive_moisture_40cm,
# 'Запаси непродуктивної вологи на глибині ґрунту 50 см, мм': self.soil_data.unproductive_moisture_50cm,
# 'Запаси непродуктивної вологи на глибині ґрунту 60 см, мм': self.soil_data.unproductive_moisture_60cm,
# 'Запаси непродуктивної вологи на глибині ґрунту 70 см, мм': self.soil_data.unproductive_moisture_70cm,
# 'Запаси непродуктивної вологи на глибині ґрунту 80 см, мм': self.soil_data.unproductive_moisture_80cm,
# 'Запаси непродуктивної вологи на глибині ґрунту 90 см, мм': self.soil_data.unproductive_moisture_90cm,
# 'Запаси непродуктивної вологи на глибині ґрунту 100 см, мм': self.soil_data.unproductive_moisture_100cm,
# 'Запаси продуктивної вологи при НВ на глибині ґрунту 10 см, мм': self.soil_data.productive_moisture_nv_10cm,
# 'Запаси продуктивної вологи при НВ на глибині ґрунту 20 см, мм': self.soil_data.productive_moisture_nv_20cm,
# 'Запаси продуктивної вологи при НВ на глибині ґрунту 30 см, мм': self.soil_data.productive_moisture_nv_30cm,
# 'Запаси продуктивної вологи при НВ на глибині ґрунту 40 см, мм': self.soil_data.productive_moisture_nv_40cm,
# 'Запаси продуктивної вологи при НВ на глибині ґрунту 50 см, мм': self.soil_data.productive_moisture_nv_50cm,
# 'Запаси продуктивної вологи при НВ на глибині ґрунту 60 см, мм': self.soil_data.productive_moisture_nv_60cm,
# 'Запаси продуктивної вологи при НВ на глибині ґрунту 70 см, мм': self.soil_data.productive_moisture_nv_70cm,
# 'Запаси продуктивної вологи при НВ на глибині ґрунту 80 см, мм': self.soil_data.productive_moisture_nv_80cm,
# 'Запаси продуктивної вологи при НВ на глибині ґрунту 90 см, мм': self.soil_data.productive_moisture_nv_90cm,
# 'Запаси продуктивної вологи при НВ на глибині ґрунту 100 см, мм': self.soil_data.productive_moisture_nv_100cm


# def add_data_1_date(self):
#     try:
#         new_row = {
#             'Рік': self.soil_data.year,
#             'Область': self.soil_data.region,
#             'Метеостанція': self.soil_data.weather_station,
#             # 'Місяць': self.soil_data.month,
#             'Дата': self.soil_data.date_1,
#             'Культура': self.soil_data.culture,
#             'Попередня культура': self.soil_data.previous_culture,
#             'Тип ґрунту': self.soil_data.soil_type,
#             'Ділянка': self.soil_data.plot,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 10 см, %': self.soil_data.moisture_content_dry_soil_avg_10cm_1_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 20 см, %': self.soil_data.moisture_content_dry_soil_avg_20cm_1_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 30 см, %': self.soil_data.moisture_content_dry_soil_avg_30cm_1_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 40 см, %': self.soil_data.moisture_content_dry_soil_avg_40cm_1_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 50 см, %': self.soil_data.moisture_content_dry_soil_avg_50cm_1_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 60 см, %': self.soil_data.moisture_content_dry_soil_avg_60cm_1_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 70 см, %': self.soil_data.moisture_content_dry_soil_avg_70cm_1_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 80 см, %': self.soil_data.moisture_content_dry_soil_avg_80cm_1_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 90 см, %': self.soil_data.moisture_content_dry_soil_avg_90cm_1_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 100 см, %': self.soil_data.moisture_content_dry_soil_avg_100cm_1_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 10 см, мм': self.soil_data.total_moisture_by_layers_10cm_1_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 20 см, мм': self.soil_data.total_moisture_by_layers_20cm_1_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 30 см, мм': self.soil_data.total_moisture_by_layers_30cm_1_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 40 см, мм': self.soil_data.total_moisture_by_layers_40cm_1_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 50 см, мм': self.soil_data.total_moisture_by_layers_50cm_1_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 60 см, мм': self.soil_data.total_moisture_by_layers_60cm_1_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 70 см, мм': self.soil_data.total_moisture_by_layers_70cm_1_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 80 см, мм': self.soil_data.total_moisture_by_layers_80cm_1_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 90 см, мм': self.soil_data.total_moisture_by_layers_90cm_1_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 100 см, мм': self.soil_data.total_moisture_by_layers_100cm_1_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 10 см, мм': self.soil_data.productive_moisture_10cm_1_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 20 см, мм': self.soil_data.productive_moisture_20cm_1_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 30 см, мм': self.soil_data.productive_moisture_30cm_1_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 40 см, мм': self.soil_data.productive_moisture_40cm_1_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 50 см, мм': self.soil_data.productive_moisture_50cm_1_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 60 см, мм': self.soil_data.productive_moisture_60cm_1_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 70 см, мм': self.soil_data.productive_moisture_70cm_1_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 80 см, мм': self.soil_data.productive_moisture_80cm_1_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 90 см, мм': self.soil_data.productive_moisture_90cm_1_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 100 см, мм': self.soil_data.productive_moisture_100cm_1_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 10 см, мм': self.soil_data.cumulative_productive_moisture_10cm_1_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 20 см, мм': self.soil_data.cumulative_productive_moisture_20cm_1_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 30 см, мм': self.soil_data.cumulative_productive_moisture_30cm_1_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 40 см, мм': self.soil_data.cumulative_productive_moisture_40cm_1_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 50 см, мм': self.soil_data.cumulative_productive_moisture_50cm_1_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 60 см, мм': self.soil_data.cumulative_productive_moisture_60cm_1_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 70 см, мм': self.soil_data.cumulative_productive_moisture_70cm_1_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 80 см, мм': self.soil_data.cumulative_productive_moisture_80cm_1_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 90 см, мм': self.soil_data.cumulative_productive_moisture_90cm_1_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 100 см, мм': self.soil_data.cumulative_productive_moisture_100cm_1_date,
#             # 'Серередня температура повітря за період спостереження, С': self.soil_data.ave_temperature_1_date,
#             # 'Сума опадів за період спостереження, мм': self.soil_data.precipitation_1_date,
#             # 'Примітка': self.soil_data.note_1_date,
#
#         }
#
#         new_row_df = pd.DataFrame([new_row])  # Створюємо DataFrame з одного рядка
#         self.data_frame = pd.concat([self.data_frame, new_row_df],
#                                     ignore_index=True)  # Додаємо рядок до основного DataFrame
#         logging.info("Data added successfully to DataFrame.")
#     except Exception as e:
#         logging.error(f"Failed to add data to DataFrame: {e}")
#
#     print("Data added successfully. Current DataFrame: ------> 1")

# def add_data_2_date(self):
#     try:
#         new_row = {
#             'Рік': self.soil_data.year,
#             'Область': self.soil_data.region,
#             'Метеостанція': self.soil_data.weather_station,
#             # 'Місяць': self.soil_data.month,
#             'Дата': self.soil_data.date_2,
#             'Культура': self.soil_data.culture,
#             'Попередня культура': self.soil_data.previous_culture,
#             'Тип ґрунту': self.soil_data.soil_type,
#             'Ділянка': self.soil_data.plot,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 10 см, %': self.soil_data.moisture_content_dry_soil_avg_10cm_2_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 20 см, %': self.soil_data.moisture_content_dry_soil_avg_20cm_2_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 30 см, %': self.soil_data.moisture_content_dry_soil_avg_30cm_2_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 40 см, %': self.soil_data.moisture_content_dry_soil_avg_40cm_2_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 50 см, %': self.soil_data.moisture_content_dry_soil_avg_50cm_2_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 60 см, %': self.soil_data.moisture_content_dry_soil_avg_60cm_2_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 70 см, %': self.soil_data.moisture_content_dry_soil_avg_70cm_2_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 80 см, %': self.soil_data.moisture_content_dry_soil_avg_80cm_2_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 90 см, %': self.soil_data.moisture_content_dry_soil_avg_90cm_2_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 100 см, %': self.soil_data.moisture_content_dry_soil_avg_100cm_2_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 10 см, мм': self.soil_data.total_moisture_by_layers_10cm_2_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 20 см, мм': self.soil_data.total_moisture_by_layers_20cm_2_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 30 см, мм': self.soil_data.total_moisture_by_layers_30cm_2_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 40 см, мм': self.soil_data.total_moisture_by_layers_40cm_2_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 50 см, мм': self.soil_data.total_moisture_by_layers_50cm_2_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 60 см, мм': self.soil_data.total_moisture_by_layers_60cm_2_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 70 см, мм': self.soil_data.total_moisture_by_layers_70cm_2_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 80 см, мм': self.soil_data.total_moisture_by_layers_80cm_2_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 90 см, мм': self.soil_data.total_moisture_by_layers_90cm_2_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 100 см, мм': self.soil_data.total_moisture_by_layers_100cm_2_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 10 см, мм': self.soil_data.productive_moisture_10cm_2_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 20 см, мм': self.soil_data.productive_moisture_20cm_2_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 30 см, мм': self.soil_data.productive_moisture_30cm_2_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 40 см, мм': self.soil_data.productive_moisture_40cm_2_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 50 см, мм': self.soil_data.productive_moisture_50cm_2_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 60 см, мм': self.soil_data.productive_moisture_60cm_2_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 70 см, мм': self.soil_data.productive_moisture_70cm_2_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 80 см, мм': self.soil_data.productive_moisture_80cm_2_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 90 см, мм': self.soil_data.productive_moisture_90cm_2_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 100 см, мм': self.soil_data.productive_moisture_100cm_2_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 10 см, мм': self.soil_data.cumulative_productive_moisture_10cm_2_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 20 см, мм': self.soil_data.cumulative_productive_moisture_20cm_2_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 30 см, мм': self.soil_data.cumulative_productive_moisture_30cm_2_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 40 см, мм': self.soil_data.cumulative_productive_moisture_40cm_2_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 50 см, мм': self.soil_data.cumulative_productive_moisture_50cm_2_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 60 см, мм': self.soil_data.cumulative_productive_moisture_60cm_2_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 70 см, мм': self.soil_data.cumulative_productive_moisture_70cm_2_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 80 см, мм': self.soil_data.cumulative_productive_moisture_80cm_2_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 90 см, мм': self.soil_data.cumulative_productive_moisture_90cm_2_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 100 см, мм': self.soil_data.cumulative_productive_moisture_100cm_2_date,
#             # 'Серередня температура повітря за період спостереження, С': self.soil_data.ave_temperature_2_date,
#             # 'Сума опадів за період спостереження, мм': self.soil_data.precipitation_2_date,
#             # 'Примітка': self.soil_data.note_2_date,
#         }
#
#         new_row_df = pd.DataFrame([new_row])  # Створюємо DataFrame з одного рядка
#         self.data_frame = pd.concat([self.data_frame, new_row_df],
#                                     ignore_index=True)  # Додаємо рядок до основного DataFrame
#         logging.info("Data added successfully to DataFrame.")
#     except Exception as e:
#         logging.error(f"Failed to add data to DataFrame: {e}")
#     print("Data added successfully. Current DataFrame:----> 2")

#
# def add_data_3_date(self):
#     try:
#         new_row = {
#             'Рік': self.soil_data.year,
#             'Область': self.soil_data.region,
#             'Метеостанція': self.soil_data.weather_station,
#             # 'Місяць': self.soil_data.month,
#             'Дата': self.soil_data.date_3,
#             'Культура': self.soil_data.culture,
#             'Попередня культура': self.soil_data.previous_culture,
#             'Тип ґрунту': self.soil_data.soil_type,
#             'Ділянка': self.soil_data.plot,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 10 см, %': self.soil_data.moisture_content_dry_soil_avg_10cm_3_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 20 см, %': self.soil_data.moisture_content_dry_soil_avg_20cm_3_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 30 см, %': self.soil_data.moisture_content_dry_soil_avg_30cm_3_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 40 см, %': self.soil_data.moisture_content_dry_soil_avg_40cm_3_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 50 см, %': self.soil_data.moisture_content_dry_soil_avg_50cm_3_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 60 см, %': self.soil_data.moisture_content_dry_soil_avg_60cm_3_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 70 см, %': self.soil_data.moisture_content_dry_soil_avg_70cm_3_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 80 см, %': self.soil_data.moisture_content_dry_soil_avg_80cm_3_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 90 см, %': self.soil_data.moisture_content_dry_soil_avg_90cm_3_date,
#             'Вологість від абсолютно сухого ґрунту (середня) на глибині ґрунту 100 см, %': self.soil_data.moisture_content_dry_soil_avg_100cm_3_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 10 см, мм': self.soil_data.total_moisture_by_layers_10cm_3_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 20 см, мм': self.soil_data.total_moisture_by_layers_20cm_3_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 30 см, мм': self.soil_data.total_moisture_by_layers_30cm_3_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 40 см, мм': self.soil_data.total_moisture_by_layers_40cm_3_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 50 см, мм': self.soil_data.total_moisture_by_layers_50cm_3_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 60 см, мм': self.soil_data.total_moisture_by_layers_60cm_3_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 70 см, мм': self.soil_data.total_moisture_by_layers_70cm_3_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 80 см, мм': self.soil_data.total_moisture_by_layers_80cm_3_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 90 см, мм': self.soil_data.total_moisture_by_layers_90cm_3_date,
#             'Запаси загальної вологи по шарах на глибині ґрунту 100 см, мм': self.soil_data.total_moisture_by_layers_100cm_3_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 10 см, мм': self.soil_data.productive_moisture_10cm_3_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 20 см, мм': self.soil_data.productive_moisture_20cm_3_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 30 см, мм': self.soil_data.productive_moisture_30cm_3_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 40 см, мм': self.soil_data.productive_moisture_40cm_3_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 50 см, мм': self.soil_data.productive_moisture_50cm_3_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 60 см, мм': self.soil_data.productive_moisture_60cm_3_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 70 см, мм': self.soil_data.productive_moisture_70cm_3_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 80 см, мм': self.soil_data.productive_moisture_80cm_3_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 90 см, мм': self.soil_data.productive_moisture_90cm_3_date,
#             'Запаси продуктивної вологи по шарах на глибині ґрунту 100 см, мм': self.soil_data.productive_moisture_100cm_3_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 10 см, мм': self.soil_data.cumulative_productive_moisture_10cm_3_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 20 см, мм': self.soil_data.cumulative_productive_moisture_20cm_3_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 30 см, мм': self.soil_data.cumulative_productive_moisture_30cm_3_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 40 см, мм': self.soil_data.cumulative_productive_moisture_40cm_3_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 50 см, мм': self.soil_data.cumulative_productive_moisture_50cm_3_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 60 см, мм': self.soil_data.cumulative_productive_moisture_60cm_1_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 70 см, мм': self.soil_data.cumulative_productive_moisture_70cm_3_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 80 см, мм': self.soil_data.cumulative_productive_moisture_80cm_3_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 90 см, мм': self.soil_data.cumulative_productive_moisture_90cm_3_date,
#             'Запаси продуктивної вологи наростаючим підсумком на глибині ґрунту 100 см, мм': self.soil_data.cumulative_productive_moisture_100cm_3_date,
#             # 'Серередня температура повітря за період спостереження, С': self.soil_data.ave_temperature_3_date,
#             # 'Сума опадів за період спостереження, мм': self.soil_data.precipitation_3_date,
#             # 'Примітка': self.soil_data.note_3_date,
#         }
#         new_row_df = pd.DataFrame([new_row])  # Створюємо DataFrame з одного рядка
#         self.data_frame = pd.concat([self.data_frame, new_row_df],
#                                     ignore_index=True)  # Додаємо рядок до основного DataFrame
#         logging.info("Data added successfully to DataFrame.")
#     except Exception as e:
#         logging.error(f"Failed to add data to DataFrame: {e}")
#     print("Data added successfully. Current DataFrame: ---> 3")
