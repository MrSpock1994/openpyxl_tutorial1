from openpyxl import Workbook, load_workbook

# Creating the workbook
wb_tutorial2 = Workbook()

# Creating the sheet
ws_tutorial2_data = wb_tutorial2.active
ws_tutorial2_data.title = "Data"

# Inserting information

ws_tutorial2_data.append(["Name", "Age", "Salary"])

# Creating loops to populate columns A, B and C

for c in range(2, 21):
    ws_tutorial2_data.append(["A", "B", "C"])


# Saving
wb_tutorial2.save('Tutorial2.xlsx')
