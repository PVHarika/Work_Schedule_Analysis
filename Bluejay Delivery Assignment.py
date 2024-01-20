pip install pandas
import pandas as pd
from datetime import datetime, timedelta

def convert_time_to_hours(time_str):
    try:
        time_format = '%m/%d/%Y %H:%M:%S'
        time_obj = datetime.strptime(time_str, time_format)
        total_hours = time_obj.hour + time_obj.minute / 60 + time_obj.second / 3600
        return total_hours
    except ValueError:
        return 0.0

def analyze_excel(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Check for the existence of required columns
    required_columns = ['Position ID', 'Position Status', 'Time', 'Time Out', 'Timecard Hours (as Time)', 'Pay Cycle End Date', 'Employee Name', 'File Number']
    for column in required_columns:
        if column not in df.columns:
            print(f"Error: Column '{column}' not found in the Excel file.")
            return

    for index, row in df.iterrows():
        position_id = row['Position ID']
        position_status = row['Position Status']

        # Check and handle missing or invalid date values
        if pd.notna(row['Time']) and pd.notna(row['Time Out']) and pd.notna(row['Pay Cycle End Date']):
            time_in = row['Time']
            time_out = row['Time Out']
            timecard_hours_str = row['Timecard Hours (as Time)']
            pay_cycle_end_date = row['Pay Cycle End Date']
            employee_name = row['Employee Name']
            file_number = row['File Number']

            # Convert 'Timecard Hours (as Time)' to total hours as a float
            timecard_hours = convert_time_to_hours(str(timecard_hours_str))

            # a) Check for employees who worked for 7 consecutive days
            consecutive_days = 0
            current_date = pd.to_datetime(pay_cycle_end_date, format='%m/%d/%Y')
            for _, next_row in df.iterrows():
                next_date = pd.to_datetime(next_row['Pay Cycle End Date'], format='%m/%d/%Y')
                if (
                    next_row['Employee Name'] == employee_name
                    and (next_date - current_date).days == 1
                ):
                    consecutive_days += 1
                    current_date = next_date
                else:
                    break

            if consecutive_days >= 6:
                print(f"{employee_name} ({position_id}) has worked for 7 consecutive days.")

            # b) Check for employees with less than 10 hours between shifts
            if consecutive_days == 0:
                next_row = df.iloc[index + 1] if index + 1 < len(df) else None
                if (next_row is not None) and pd.notna(next_row['Time']):
                    next_time_in = next_row['Time']
                    time_between_shifts = (next_time_in - time_out).total_seconds() / 3600
                    if 1 < time_between_shifts < 10:
                        print(f"{employee_name} ({position_id}) has less than 10 hours between shifts.")

            # c) Check for employees who worked for more than 14 hours in a single shift
            if timecard_hours > 14:
                print(f"{employee_name} ({position_id}) has worked for more than 14 hours in a single shift.")
        else:
            print(f"Skipping row {index + 1} due to missing or invalid date values.")

# Example usage
file_path = '/content/Assignment_Timecard.xlsx'
analyze_excel(file_path)

