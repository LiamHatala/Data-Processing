import pandas as pd
import os
from datetime import datetime

# Define file paths
input_csv = r"" # This is the file path of your input excel file
output_excel = "" # This is the Name of the file that this program will output

# Define relevant columns
columns_to_use = ["Row #", "LOCAL_TIME", "SOURCE", "CONDITIONNAME", "ACTION"]

# Load CSV file
df = pd.read_csv(input_csv, usecols=columns_to_use, delimiter=',')

# Sort by "Row #" in ascending order (1 → highest) so it's processed in correct time sequence
df = df.sort_values(by="Row #", ascending=True).reset_index(drop=True)

# Convert LOCAL_TIME to datetime
df["LOCAL_TIME"] = pd.to_datetime(df["LOCAL_TIME"], errors="coerce")

# Normalize column values
df["ACTION"] = df["ACTION"].replace("", None).fillna("").astype(str).str.strip().str.lower()
df["CONDITIONNAME"] = df["CONDITIONNAME"].fillna("").astype(str).str.strip().str.upper()

# Debugging: Print first 10 rows to confirm `Row #` column is correct
print("\nFirst 10 rows after sorting (checking `Row #` column consistency):")
print(df.head(10))

# Initialize storage for alarm tracking
alarm_records = []
active_alarms = {}

# Process each row sequentially
for index, row in df.iterrows():
    print(f"\nProcessing row {row['Row #']} (Excel Row {index+2}): {row['LOCAL_TIME']} - {row['SOURCE']} - {row['CONDITIONNAME']} - {row['ACTION']}")

    alarm_key = (row["SOURCE"], row["CONDITIONNAME"])  # Unique identifier for each alarm
    action = row["ACTION"]
    time = row["LOCAL_TIME"]
    row_number = row["Row #"]  # Ensure we are using the actual Row # from the file

    # Alarm End Condition: If an OK is found, close the existing alarm
    if action == "ok" and alarm_key in active_alarms:
        start_data = active_alarms.pop(alarm_key)
        start_time = start_data["time"]
        start_row = start_data["row"]

        if time > start_time:  # Ensure non-zero time difference
            duration = time - start_time
            duration_seconds = int(duration.total_seconds())

            alarm_records.append([
                start_row, row_number, row["SOURCE"], row["CONDITIONNAME"],
                f"{start_time.strftime('%I:%M:%S %p')} (Row {start_row})",
                f"{time.strftime('%I:%M:%S %p')} (Row {row_number})",
                str(duration), duration_seconds
            ])
            print(f"End Alarm: {row['SOURCE']} at {time} (Row {row_number}) - Duration: {duration}")

    # Alarm Start Condition: Start if it's a FAIL or STEPTO with an empty ACTION
    elif row["CONDITIONNAME"] in ["FAIL", "STEPTO"] and row["ACTION"] == "":
        print(f"Start Alarm: {row['SOURCE']} at {time} (Row {row_number})")
        active_alarms[alarm_key] = {"time": time, "row": row_number}

# Debugging: Check how many alarms were processed
print(f"\nTotal alarms processed: {len(alarm_records)}")

# Convert to DataFrame
processed_df = pd.DataFrame(alarm_records, columns=[
    "Start Row #", "OK Row #", "Device_Name", "Condition",
    "Alarm_Time (Row #)", "OK_Time (Row #)", "Time_Taken", "Time_Taken_Seconds"
])

# Ensure Excel order follows chronological order
processed_df = processed_df.sort_values(by="Start Row #", ascending=True).reset_index(drop=True)

# Check if DataFrame is empty
if processed_df.empty:
    print("\nNo alarms were processed! Check your input data and filters.")
else:
    print(f"\nSaving {len(processed_df)} alarms to {output_excel}")
    print(processed_df.head())  # Print first few rows to confirm

# Remove old Excel file before writing to prevent overwrite issues
if os.path.exists(output_excel):
    os.remove(output_excel)
    print(f"Removed old file: {output_excel}")

# Save to Excel
with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
    processed_df.to_excel(writer, index=False, sheet_name="Sheet1")

print(f"Processed alarm data saved to {output_excel}")
print(f"heck the output file here: {os.path.abspath(output_excel)}")
