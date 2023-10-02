import pandas as pd

# Load the Excel file into a DataFrame
file_path = 'employee_data.xlsx'  
df = pd.read_excel(file_path)

# Columns: ['Position ID', 'Position Status', 'Time In', 'Time Out', 'Timecard Hours (as Time)', 'Pay Cycle Start Date', 'Pay Cycle End Date', 'Employee Name', 'File Number']

# Convert Time In and Time Out columns to datetime objects
df['Time'] = pd.to_datetime(df['Time'])
df['Time Out'] = pd.to_datetime(df['Time Out'])


df.sort_values(by=['Employee Name', 'Time'], inplace=True)

# Initialize variables for tracking consecutive days and previous shift end time
consecutive_days = 1
prev_shift_end = None

# Initialize a list to store the employees who meet the criteria
employees_consecutive_7_days = []
employees_less_than_10_hours = []
employees_more_than_14_hours = []

# Iterate through the DataFrame to analyze shifts
for index, row in df.iterrows():
    employee_name = row['Employee Name']
    time_in = row['Time']
    time_out = row['Time Out']

    # Check for consecutive days
    if prev_shift_end is not None:
        if (time_in - prev_shift_end).total_seconds() >= 86400:  # 86400 seconds in a day
            consecutive_days = 1
        else:
            consecutive_days += 1

    # Check for shifts with less than 10 hours between
    if prev_shift_end is not None:
        time_between_shifts = (time_in - prev_shift_end).total_seconds() / 3600  # Convert to hours
        if 1 < time_between_shifts < 10:
            employees_less_than_10_hours.append(employee_name)

    # Check for shifts with more than 14 hours
    shift_duration = (time_out - time_in).total_seconds() / 3600  # Convert to hours
    if shift_duration > 14:
        employees_more_than_14_hours.append(employee_name)

    # Update previous shift end time
    prev_shift_end = time_out

    # Check for 7 consecutive days
    if consecutive_days >= 7:
        if employee_name not in employees_consecutive_7_days:
            employees_consecutive_7_days.append(employee_name)


print("Employees who worked for 7 consecutive days:")
for employee in employees_consecutive_7_days:
    print(employee)

print("\nEmployees with less than 10 hours between shifts:")
for employee in employees_less_than_10_hours:
    print(employee)

print("\nEmployees who worked for more than 14 hours in a single shift:")
for employee in employees_more_than_14_hours:
    print(employee)

# Save the results to output.txt
with open('output.txt', 'w') as output_file:
    output_file.write("Employees who worked for 7 consecutive days:\n")
    for employee in employees_consecutive_7_days:
        output_file.write(employee + '\n')

    output_file.write("\nEmployees with less than 10 hours between shifts:\n")
    for employee in employees_less_than_10_hours:
        output_file.write(employee + '\n')

    output_file.write("\nEmployees who worked for more than 14 hours in a single shift:\n")
    for employee in employees_more_than_14_hours:
        output_file.write(employee + '\n')
