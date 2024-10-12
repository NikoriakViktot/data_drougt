import os
import re

def remove_duplicate_years(filename):
    # Enhanced regex to catch consecutive duplicate years spread across the filename
    # This finds patterns of a year repeated exactly after itself and removes the duplicate.
    filename = re.sub(r'(\b\d{4}\b)(_\1)+', r'\1', filename)
    # Normalize underscores in filenames
    filename = re.sub(r'__+', '_', filename)  # Replace multiple underscores with a single one
    filename = filename.strip('_')  # Trim leading and trailing underscores
    return filename

def rename_files_in_directory(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            new_name = remove_duplicate_years(file)
            old_file_path = os.path.join(root, file)
            new_file_path = os.path.join(root, new_name)
            if new_name != file:
                # Ensure that the new filename does not already exist in the directory
                if not os.path.exists(new_file_path):
                    os.rename(old_file_path, new_file_path)
                    print(f"Renamed '{old_file_path}' to '{new_file_path}'")
                else:
                    print(f"Cannot rename '{old_file_path}' because '{new_file_path}' already exists.")

# Specify the path to the directory where the files are located
directory_path = r"C:\Users\user\PycharmProjects\data_drought\DATA_base_soil_water"
rename_files_in_directory(directory_path)
