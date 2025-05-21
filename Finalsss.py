from openpyxl import Workbook  

# Create a new Excel workbook
workbook = Workbook()
sheet = workbook.active  # Get the active sheet
sheet.title = "Grades"

# Write data into cells
sheet["A1"] = "Name"
sheet["B1"] = "Course"
sheet["C1"] = "Grade"

# Save the workbook
workbook.save("Grade.xlsx")
print("Excel file created successfully!")