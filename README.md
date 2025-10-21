🧙 Expense Wizard - Spending Report Generator
---------------------------------------------

Version: 1.0  
Author: DG3030  
Website: [NULL]

**Backstory**

When I owned a jewelry shop, one of the most time-consuming tasks at the end of every month was handling accounts payable. This meant cross-referencing every credit card payment a tedious, error-prone process when done by hand.

After returning to school and learning Python, one of my first goals was to find a better way to handle this problem. I started with Python’s built-in libraries but quickly realized they weren’t enough — that’s when Pandas came into the picture. With guidance from our accountant on what information was most relevant, the application began to take shape.

The result was a tool that automatically reads credit card statements, sorts and categorizes each transaction using the item codes defined by the credit card company, and generates visual summaries. It turns hours of manual sorting into minutes — and because accountants apparently hate pie charts, that’s the default visualization.


Overview:
---------
Expense Wizard helps you quickly parse trhough a credit card statements, organize them by date-filtering and ceating an expense report.  
It supports Excel and CSV formats and offers a variety of chart styles to help visualize where your money goes.

Requirements:
---------
To run Expense Wizard, you’ll need:

- Python 3.x
- pandas
- PyQt5
- openpyxl

You can install the required packages with:

pip install -r requirements.txt

Features:
---------
 Import multiple statement files (.xlsx format)  
 Filter by custom date ranges  
 Group expenses by Weekly, Biweekly, or Monthly  
 Export to:
   - Excel (.xlsx) with summary + category sheets
   - CSV (.csv) split by time period  
 Choose chart types: Pie, Bar, Column, Doughnut, Radar  
 Automatically remembers your last-used settings  

How to Use:
-----------
1. Launch the app  
2. Click "Browse" next to Input Folder and select the folder containing your credit card statements  
   (Make sure you’ve saved all your .xlsx statements locally in one folder)  
3. Set a save filename and location for your output  
4. Pick your date range using the calendar  
5. Choose your grouping method (Weekly, Biweekly, Monthly)  
6. Choose your export format (Excel or CSV)  
7. Select a chart type (Pie, Bar, Column, Doughnut, Radar, Treemap*)  
8. Click "Generate Report"  

Your report will be saved to the location you specified.

Note:
-----
- CSV exports create a separate file for each time period  
- Excel reports include a summary sheet, category sheets, and visual charts

ExpenseWizrd.exe file
-----
The executable file does not have any requirement... obviously its all packaged and ready to roll.
  

Uninstall Instructions:
-----------------------
1. Open Windows Settings > Apps  
2. Search for "Expense Wizard"  
3. Click "Uninstall"  

Support:
--------
There is none. Good luck. 

Thanks for using Expense Wizard!
