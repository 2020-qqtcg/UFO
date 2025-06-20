import pandas as pd
import os
import shutil

# from io import StringIO # This import is no longer needed

# --- IMPORTANT: Please update this path to your actual Excel file location ---
excel_file_path = r"C:\Users\v-yuhangxie\UFO\benchmark\sheetcoplit\SheetCopilot\dataset\dataset.xlsx"
# -----------------------------------------------------------------------------

try:
    # Read the Excel data from the specified file
    # Assuming the data is in the first sheet or a sheet named 'Sheet1'
    df = pd.read_excel(excel_file_path)
    print(f"Successfully loaded data from: {excel_file_path}")
except FileNotFoundError:
    print(f"Error: Excel file not found at '{excel_file_path}'. Please check the path.")
    exit()  # Exit the script if the file isn't found
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()  # Exit for other reading errors

# Define the base directory where your source folders are located
# IMPORTANT: Replace this with the actual path if it's different on your system.
source_base_dir = r"C:\Users\v-yuhangxie\UFO\logs\20250530_copilot"

# Define the new base directory where categorized folders will be created
# This script will create this directory if it doesn't exist.
destination_base_dir = r"C:\Users\v-yuhangxie\UFO\logs\20250530_copilot_categorized"

# Create the destination base directory if it doesn't exist
os.makedirs(destination_base_dir, exist_ok=True)
print(f"Ensured destination base directory exists: {destination_base_dir}")

# Process each row in the DataFrame
for index, row in df.iterrows():
    # Construct the name of the source folder (e.g., "1_BoomerangSales")
    # Ensure 'No.' and 'Sheet Name' columns exist and are not NaN
    if 'No.' not in row or pd.isna(row['No.']):
        print(f"Skipping row {index}: 'No.' column missing or invalid.")
        continue
    if 'Sheet Name' not in row or pd.isna(row['Sheet Name']):
        print(f"Skipping row {index}: 'Sheet Name' column missing or invalid.")
        continue

    folder_no = int(row['No.'])  # Convert to integer as it's typically a number
    sheet_name = str(row['Sheet Name']).strip()  # Ensure it's a string and strip whitespace
    source_folder_name = f"{folder_no}_{sheet_name}"
    full_source_path = os.path.join(source_base_dir, source_folder_name)

    # Get the categories for the current folder, clean up whitespace, and split into a list
    categories_str = str(row['Categories']).strip()  # Ensure it's a string and remove leading/trailing whitespace

    # Handle cases where Categories might be NaN or empty after stripping
    if not categories_str or categories_str.lower() == 'nan':
        print(f"Skipping '{source_folder_name}': No valid categories found.")
        continue  # Skip to the next folder

    # Split categories by comma and strip whitespace from each individual category
    categories = [cat.strip() for cat in categories_str.split(',')]

    # Check if the source folder actually exists before trying to copy
    if not os.path.exists(full_source_path):
        print(f"Warning: Source folder '{full_source_path}' does not exist. Skipping.")
        continue

    # For each category, create the destination folder and copy the source folder into it
    for category in categories:
        # Sanitize category name for use as a folder name (remove invalid characters if any)
        # For simplicity, this example assumes category names are valid for folder names.
        # If your categories might contain characters like /, \, :, *, ?, ", <, >, |, you'd need to sanitize them.

        # Create the full path for the category-specific destination
        category_destination_path = os.path.join(destination_base_dir, category)

        # Create the category-specific destination directory if it doesn't exist
        os.makedirs(category_destination_path, exist_ok=True)
        print(f"Ensured category directory exists: {category_destination_path}")

        # Construct the final destination path for the copied folder
        final_destination_path = os.path.join(category_destination_path, source_folder_name)

        # Copy the source folder to the new category directory
        # If the destination already exists, shutil.copytree will raise an error.
        # So, we remove it first if it exists to allow overwriting (or handle differently based on requirement)
        if os.path.exists(final_destination_path):
            print(f"Destination '{final_destination_path}' already exists. Removing before copying.")
            shutil.rmtree(final_destination_path)

        try:
            shutil.copytree(full_source_path, final_destination_path)
            print(f"Copied '{source_folder_name}' to '{category_destination_path}'")
        except shutil.Error as e:
            print(f"Error copying '{source_folder_name}' to '{category_destination_path}': {e}")
        except OSError as e:
            print(f"OS Error while copying '{source_folder_name}' to '{category_destination_path}': {e}")

print("\nFolder organization complete!")
print(f"Check the '{destination_base_dir}' directory for the categorized folders.")
