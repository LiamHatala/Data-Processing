import os
import csv
import glob
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


# Add the folder path for where all of the channel .cvs files are
folder_path = r"Folder_Path"
output_file = os.path.join(folder_path, "CombinedTagNames.xlsx")

# Find all CSV files in your folder
csv_files = glob.glob(os.path.join(folder_path, "*.csv"))

# We'll accumulate partial DataFrames here
all_data = []

for file_path in csv_files:
    # -- 1) Extract the top-left cell from the first line:
    with open(file_path, mode='r', newline='', encoding='utf-8') as f:
        reader = csv.reader(f, delimiter=',')  # Use delimiter=';' if semicolon, etc.
        first_row = next(reader, [])
        top_left_name = first_row[0] if first_row else ""
    
    # -- 2) Read the CSV with Pandas, skipping lines so that "TagName" is your header:
    #    If the CSV has the header on line 3, skiprows=2. If it's on line 2, skiprows=1, etc.
    df = pd.read_csv(file_path, delimiter=',', encoding='utf-8', skiprows=2)
    
    # Make sure the dataframe *has* a TagName column:
    if "Tag Name" not in df.columns:
        print(f"WARNING: No 'Tag Name' column in {os.path.basename(file_path)}. Skipping.")
        continue
    
    # -- 3) Create a smaller DataFrame with just TagName + an added "SystemName" column
    subset = df[["Tag Name"]].copy()
    subset.insert(0, "SystemName", top_left_name)  # Insert to the left of TagName
    
    # Store in our master list
    all_data.append(subset)

# If we found nothing, just quit
if not all_data:
    print("No valid CSV data found (or no TagName columns). Exiting.")
    raise SystemExit

# -- 4) Combine all subsets into one final DataFrame
combined_df = pd.concat(all_data, ignore_index=True)

# -- 5) Write everything to a single Excel sheet
wb = Workbook()
ws = wb.active
ws.title = "AllTagNames"

# Convert DataFrame to rows
for i, row in enumerate(dataframe_to_rows(combined_df, index=False, header=True)):
    ws.append(row)

wb.save(output_file)
print(f"Success! Created {output_file}")
