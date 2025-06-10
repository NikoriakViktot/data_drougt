import os
import pandas as pd

# Define the root directory containing the data
path_all_folders = r"C:\Users\5302\PycharmProjects\data_drougt\DATA_base_soil_water"
year = '2021'  # Specify the year to process

# Define the path to the year folder
year_path = os.path.join(path_all_folders, f'ТСГ-6_{year}')

# Initialize a list to store the results
results = []

# Check if the year folder exists
if os.path.exists(year_path):
    for folder_obl in os.listdir(year_path):
        folder_obl_path = os.path.join(year_path, folder_obl)

        for folder_st in os.listdir(folder_obl_path):
            folder_st_path = os.path.join(folder_obl_path, folder_st)

            # Count the number of files in the station directory
            file_count = sum([len(files) for _, _, files in os.walk(folder_st_path)])

            # Store the data in the results list
            results.append({
                'Область': folder_obl[11:],  # Extract the region name from folder
                'Метеостанція': folder_st,
                'Кількість файлів': file_count
            })

# Convert the results list to a DataFrame
df_results = pd.DataFrame(results)
total_stations = df_results['Метеостанція'].nunique()

# Load the processed CSV file
df_results = pd.read_csv('meteostations_summary_2021.csv')

# Calculate the total number of files across all meteorological stations
total_files = df_results['Кількість файлів'].sum()
m_total_files = df_results['Кількість файлів'].mean()

# Display the total number of files
print(f"Загальна кількість метеостанцій: {total_stations}")
print(f"Загальна кількість файлів: {total_files }")
print(f"Загальна частота приблизно два спостереження у файлі: {total_files * 2}")

print(f"Загальна частота приблизно два спостереження у файлі: {m_total_files*2}")

# Save the results to a CSV file
output_csv = 'meteostations_summary_2021.csv'
df_results.to_csv(output_csv, index=False, encoding='utf-8-sig')

print(f"Processing completed. Results saved to {output_csv}")

# Display the total count of stations

print(total_stations)
# Display the first few rows of the results
print(df_results.head())
