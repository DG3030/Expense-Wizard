🧙 Expense Wizard - Spending Report Generator
---------------------------------------------

Version: 1.0  
Author: DG3030  
Website: [pff]

Back story:
---------

Between my wife and me, she handles most of our expenses, given her background, it makes sense. Occasionally she asks me to help out, but the thought of manually parsing statements, cutting and pasting everything into an Excel file? No way.  

So instead, I turned to Python, Pandas, OpenPyXL, and a few charts to cut through the manual labor. This tool reads through our credit card statements, separates each expense by category, and shows us exactly where the money is going and because aperently accountants hate pie charts, its the default setting.


Overview:
---------
Expense Wizard helps you quickly parse trhough a credit card statements, organize them by date-filtering and ceating an expense report.  
It supports Excel and CSV formats and offers a variety of chart styles to help visualize where your money goes.

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
  

Uninstall Instructions:
-----------------------
1. Open Windows Settings > Apps  
2. Search for "Expense Wizard"  
3. Click "Uninstall"  

Support:
--------
There is none. Good luck. 

Thanks for using Expense Wizard!
