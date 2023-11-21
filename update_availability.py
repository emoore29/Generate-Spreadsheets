import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

def update_availability_template(template_path, output_folder, new_date):
    # load availability template
    wb = load_workbook(template_path)
    
    # open the currently active sheet in the workbook
    ws = wb.active
      
    # Get the merged cell range
    merged_cells = ws.merged_cells.ranges
    merged_range = None
    for merged_cell in merged_cells:
        if ws['B4'] in merged_cell:
            merged_range = merged_cell
            break

    # Update the values in the individual cells of the merged cell
    for cell in merged_range:
        cell.value = new_date

    # Save the updated template to a new Excel file
    new_file_name = f"Emmaline Moore - Availability Week Commencing {new_date.strftime('%d %B %Y')}.xlsx"
    new_file_path = f"{output_folder}/{new_file_name}"
    wb.save(new_file_path)
    

# run the following code if update_availability.py is run as the "main" program and not imported into another script
if __name__ == "__main__":
    template_path = "C:/Users/emoor/OneDrive/Documents/VIQ/Availability/Emmaline Moore - Availability Week Commencing 27 November 2023.xlsx"
    output_folder = "C:/Users/emoor/OneDrive/Documents/VIQ/Availability"
    
    new_date_input = input("Enter the new date (YYYY-MM-DD): ")
    new_date = datetime.strptime(new_date_input, "%Y-%m-%d")

    update_availability_template(template_path, output_folder, new_date)