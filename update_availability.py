from datetime import datetime, timedelta
from openpyxl import load_workbook


def generate_sheets(date, number_sheets):
    dates = []
    for _ in range(number_sheets):
        dates.append(date)
        # define date, template path, output folder
        template_path = "C:/Users/emoor/OneDrive/Documents/VIQ/Availability/Emmaline Moore - Availability Template.xlsx"
        output_folder = "C:/Users/emoor/OneDrive/Documents/VIQ/Availability"

        # load availability template
        wb = load_workbook(template_path)
            
        # open the currently active sheet in the workbook
        sheet = wb.worksheets[0]

        # Update date in cell B4
        new_start_date = sheet["B4"].value=date

        # Save the updated template to a new Excel file
        new_file_name = f"Emmaline Moore - Availability Week Commencing {date.strftime('%d %B %Y')}.xlsx"
        new_file_path = f"{output_folder}/{new_file_name}"
        wb.save(new_file_path)
        
        # Increment to next Monday
        date += timedelta(days=7)
        
    formatted_dates = [d.strftime('%d %B %Y') for d in dates]
    print(f"New availability spreadsheets were created for the following dates: {formatted_dates}")


# Print welcome message
print("Hello! Welcome to Auto Availability!")

# Get the current date
current_date = datetime.now().date()

# find the date of next Monday
days_until_monday = (7 - current_date.weekday()) % 7
next_date = current_date + timedelta(days=days_until_monday)

# Welcome msg How many weeks of availability do you want to generate, starting from next Monday
start_next_monday = input(f"Today is {current_date}. Next Monday is {next_date}. Do you want to generate sheets starting next Monday? y/n: ")

if start_next_monday == "y":
    number_sheets = int(input("How many sheets do you want to generate? "))
    generate_sheets(next_date, number_sheets)
elif start_next_monday == "n":
    alt_start_date = input("What date would you like to start from? Make sure it's a Monday in this format: YYYY-MM-DD: ")
    alt_start_date = datetime.strptime(alt_start_date, "%Y-%m-%d")
    number_sheets = int(input("How many sheets do you want to generate? "))
    generate_sheets(alt_start_date, number_sheets)
else:
    print("Please restart, you need to input either 'y' or 'n' for the program to work.")






