"""
excel_monthly_report.py is a Python script that reads .xlsx files from the data folder, 
sums current (A) values per file and overall, and exports a summary report to summary.xlsx
"""

# --- Imports ---
from pathlib import Path
from openpyxl import load_workbook, Workbook

from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

# --- Configuration & Initialization ---
folder = Path("data")
files = sorted(folder.glob("*.xlsx"))

if not folder.exists():
    print("Folder not found!")
    exit()

if len(files) == 0:
    print(f"No Excel files found in {folder} folder")
    exit()


files_open = 0
files_read = 0
grand_total = 0

report_wb_out = Workbook()
report_ws_out = report_wb_out.active
report_ws_out.append(["Month", "Total_A"])

# --- Main Program ---
#Process each .xlsx file, skip temporary (~$) files, and sum current per file and overall.
for file in files:  
    if file.name.startswith("~$"):
        files_open += 1
        continue
    files_read += 1
        
    ws = load_workbook(file).active
    month_total = 0

    #Add non-empty current values to the monthly and grand totals.
    for name, current in ws.iter_rows(min_row = 2, values_only = True): 
        if current is None:
            continue
        
        if isinstance(current, str):
            cleaned = current.strip().replace(",", ".")
            if cleaned == "":
                continue
            current = float(cleaned)
        
        month_total += current
        grand_total += current
        
    # Show and save the month total.  
    print(f"{file.name}: ---> {month_total:.1f} A")
    report_ws_out.append([file.stem, round(month_total, 1)])

# --- Terminal Summary --- 
print(f"Grand total: {grand_total:.1f} A")

if files_open == 1:
    print(files_open, "File is open in EXCEL")
elif files_open == 0:
    pass
else: 
    print(files_open, "Files are open in EXCEL")

print("Files read:", files_read)

# --- Chart ---
chart = BarChart()

chart.y_axis.scaling.min = 0
chart.varyColors = False
chart.legend = None


# Which cells (row 1 is the header, so the last month is files_read + 1)
last_row = files_read + 1    # row of the last month (row 1 is the header)
data = Reference(report_ws_out, min_col=2, min_row=1, max_row=last_row)
cats = Reference(report_ws_out, min_col=1, min_row=2, max_row=last_row)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

chart.title = "Total current per month"
chart.x_axis.title = None
chart.y_axis.title = "Current (A)"

chart.x_axis.delete = False
chart.y_axis.delete = False

chart.title.overlay = False
chart.y_axis.title.overlay = False

chart.width = 18
chart.height = 9

chart.dataLabels = DataLabelList()
chart.dataLabels.showVal = True
chart.dataLabels.showSerName = False     
chart.dataLabels.showCatName = False     
chart.dataLabels.showLegendKey = False

report_ws_out.add_chart(chart, "D2")


#--- Export to Excel ---
report_ws_out.append([])
report_ws_out.append(["GRAND TOTAL current(A):", round(grand_total, 1)])
report_ws_out.append(["FILES READ:", files_read])

report_ws_out.column_dimensions["A"].width = 25

report_wb_out.save("summary.xlsx")
print("📊 Excel Report saved to: summary.xlsx")