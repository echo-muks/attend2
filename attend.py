import openpyxl
from datetime import datetime, timedelta

def record_attendance(employee_name, daily_working_hours=8, allowed_delay=15):
    excel_file = "attendance_record.xlsx"
    
    try:
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Employee Name", "Attendance Time", "Month", "Number of Attendance Days", "Delay Hours", "Notes"])

    attendance_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    registration_month = datetime.now().strftime("%B")

    # Calculate the number of days and delay hours
    today_weekday = datetime.now().weekday()
    if today_weekday < 5:  # If it's a working day
        delay_hours = allowed_delay
    else:
        delay_hours = 0

    sheet.append([employee_name, attendance_time, registration_month, None, delay_hours, ""])

    workbook.save(excel_file)

# You can modify the employee name and daily working hours
record_attendance("Employee Name", daily_working_hours=8, allowed_delay=15)
