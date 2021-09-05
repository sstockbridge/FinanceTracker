# FinanceTracker
Finance tracking desktop application made in python using PySimpleGUI, OpenPyXL, and MatPlotLib. 
The application uses an .xlsx file to track finances and will save one onto your computer (in the same directory) upon opening the application.

How to Start:
1) Download all files into .zip folder, then extract into desired directory.
2) Run main.exe and enjoy. Make sure that main.exe and tracking_sheet.py are in the same directory. Note that .xlsx file will be saved onto the same directory as main.exe.

How to Use:
1) Must fill out all fields before adding earned income or an expense. "Add" will also not work if the "Amount" field does not contain a numerical value. The other fields can contain alphanumeric numbers or symbols.
2) NOTE: Do not enter a "$" symbol into any field, it causes an issue where the "Income/Expense Summary" does not display anything if entered. If entered, simply change the character in the .xlsx sheet.
3) The "Amount" field works by adding a "-" sign before the value if it is an expense. If it is income earned, no sign is needed. For example, if desired value is an expense of $100, simply input "-100" without the quotes. If earned $100, simply put "100". Decimal values function the same way.
4) "Reset Tracker" deletes everything off of your .xlsx file. If the .xlsx file is removed from the directory, the program will simply create another one automatically.
5) NOTE: If you want to start a fresh sheet, simply take the existing one out of the directory and restart the program. OR click "Reset Tracker" after taking the sheet out of the directory.
6) The date displayed on the program is the date the .xlsx sheet was created. Reseting the tracker also resets this date to current date.
7) Enjoy!

