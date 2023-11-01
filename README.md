# VBA-challenge
Week 2 of the Data Analysis Boot Camp

### Introduction
This project uses VBA to take stock ticker data from an existing spreadsheet, analyze the data, and summarizes the spreadsheet's results. The final step is saving the results to an output directory so that the original data remains untouched.

The analysis performed is:
* Calculate the yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
* Calculate the percent change from opening price at the beginning of a given year to the closing price at the end of that year.
* Sum the total stock volume of the stock.
After completing the analysis, the VBA code changes the "Yearly Change" cells background color:
* Values greater than or equal to zero are colored green
* Values less than zero are colored red

Further anlysis that is performed is:
* Determine the greatest percent increase
* Determine the greatest percent decrease
* Determine the greatest total volume

### Usage
Within the master VBA macro filefile, there are three worksheets:
* Lookup: this sheet holds the file names and the legal worksheets for the resources data. This data gets populated in the drop-down menus on the other worksheets.
* SingleSheet: this sheet has two drop-down menus (these select the workbook and worksheet to analyze) and a "Run the Single Script" button. The VBA code will only analyze the data on a single sheet. SingleSheet is a code testing sheet.
* MultiSheet: this sheet has one drop-down menu (to select the workbook to analyze) and a "Run the Multi Script" button. The VBA code analyzes all of the sheets in the workbook. The multi-script is a production-quality version of the VBA code.

The MultiSheet option is used. Upon choosing the workbook, click on the "Run the Multi Script" button. The VBA code opens the workbook, runs the analysis on all worksheets, then saves it to the Output directory.
