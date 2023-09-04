import xlwings as xw

# Open the Excel file
file_path = 'مبيعات20-8-2023.xlsx'
wb = xw.Book(file_path)

# Choose the sheet you want to modify
sheet_name = 'Sales Log Ar'
sheet = wb.sheets[sheet_name]

# Set horizontal centering
sheet.api.PageSetup.CenterHorizontally = True

# Set "Fit to Page" option
sheet.api.PageSetup.Zoom = False
sheet.api.PageSetup.FitToPagesWide = 1
sheet.api.PageSetup.FitToPagesTall = 1

# Save the modified Excel file
wb.save()
wb.close()
