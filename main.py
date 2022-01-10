import pandas as pd
import sqlite3
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font

save_path = "manLog.xlsx"

if len(sys.argv) != 3:
    print("Usage: python main.py [db_path] [date_filter}\n\tdb_path: Path to sqlite3 database\n\tdate_filter: Year of repair work");
    quit()
else:
    db_path = sys.argv[1]
    date_filter = sys.argv[2]

connection = sqlite3.connect(db_path)

query = "SELECT * FROM MaintenanceLog;"

headers = ["PropCode", "Unit", "Date", "Maintenance Work", "Hours", "Details"]

df = pd.read_sql_query(query, connection)

df.drop("repair_ISO8601", axis = "columns", inplace = True)

df.rename(columns = {
    "property_code":"PropCode",
    "unit":"Unit",
    "repair_date":"Date",
    "maintenance_details":"Maintenance Work",
    "hours":"Hours",
    "details":"Details"
    #"repair_date":"Date"
    }, inplace = True)

df = df[df.Date.str.contains(date_filter)]

#pull hours column

str_hours = df.Hours
str_hours.dropna(inplace = True)

hours = 0
min = 0

for session in str_hours:
    try:
        hours += int(session[0:2])
    except:
        try:
            hours += int(session[0:1])
        except:
            print("Incorrectly formatted hours, expected [HH:MM] or [H:MM]")

    min += int(session[3:5])

hours += int(min/60)
min -= int(min/60) * 60

total_time = str(hours) + ":" + str(min)

df.to_excel(save_path, index = False)

workbook = load_workbook(filename = save_path)
sheet = workbook.active

used_cols = len(list(sheet.columns)[1]) + 1
sheet["E" + str(used_cols)].font = Font(bold = True)
sheet["D" + str(used_cols)].font = Font(bold = True)
sheet["E" + str(used_cols)] = total_time
sheet["D" + str(used_cols)] = "Total"

workbook.save(save_path)
