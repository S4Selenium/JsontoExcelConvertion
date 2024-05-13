import os
import json
import win32com.client as win32
"""
Step 1 . Read the JSON file
"""
json_data = json.loads(open('testdata.json').read())
print(json_data)

"""
Step 2 Examining the data and flatten the records into a 2D layout
"""
rows =[]

for record in json_data:
    ID=record['id']
    AdminId = record['admin_id']
    name=record['name']
    rows.append([ID,AdminId,name])

    """Step 3 Inserting records to excel spread sheet"""
    ExcelApp =win32.Dispatch('Excel.Application')
    ExcelApp.visible = True

    wb=ExcelApp.Workbooks.Add()
    ws=wb.Worksheets(1)

header_labels =('ID','AdminId','name')

#insert header labels
for indx, val in enumerate(header_labels):
    ws.Cells(1, indx + 1).Value= val

#insert Records
row_tracker = 2 
column_size = len(header_labels)

for row in rows:
    ws.Range(
        ws.Cells(row_tracker,1),
        ws.Cells(row_tracker,column_size)
    ).value = row
    row_tracker +=1