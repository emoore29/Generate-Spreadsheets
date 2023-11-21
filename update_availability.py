import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

def update_availability_template(template_path, output_folder, new_date):
    # load availability template
    wb = load_workbook(template_path)
    
    # open the currently active sheet in the workbook
    ws = wb.active
      
    # update cell B4 with new date
    ws["B4"] = new_date
    ws["C4"] = new_date


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