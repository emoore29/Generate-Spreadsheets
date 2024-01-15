# Generate Spreadsheets

This is a small program I wrote to automate a tedious work task of updating a weekly availability spreadsheet for sending my weekly availability to the company I work for. In the availability template, a user can set their availability for each day of the week, starting on a given date on a Monday.

My program opens the template and updates the date in the "B4" cell (i.e. the Monday), which then autopopulates dates for the remaining days of the week. It prompts the user to decide on a "start date" for the spreadsheets being created, whether that be next Monday or a Monday in three weeks. Then it will autogenerate the number of spreadsheets requested, starting on the start date given, with each new spreadsheet starting on the following consecutive Monday.

Currently it doesn't allow the user to change anything else in the spreadsheet, but in my case I'm not often updating the actual content, simply the date, so the program serves its purpose.

## How to run

In Anaconda Prompt:

cd to location of py file, e.g.: C:\Users\[USERNAME]\OneDrive\Documents\GitHub\Generate-Spreadsheets  

Run the following:  
python update_availability.py

Then follow the prompts. 

(This program only works with the availability template in a specific folder. If you want to use it you will need to create your own availability template and update the template_path.)


## Example:

Today is 21/11/23. Next Monday is 27/11/23. If the user wants 4 spreadsheets generated starting on next Monday, the program will create 4 spreadsheets, with the B4 Monday cell of each spreadsheet being the following:

27/11/23, 04/12/23, 11/12/23, and 18/12/23.
