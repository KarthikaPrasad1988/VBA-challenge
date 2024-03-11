# VBA-challenge 
VBA-Challenge to make a Stock Market Analysis.

## Background
This project is a VBA scripting to analyze real stock market data.This analysis used two data-points one is the [Test Data ](./VBA_Alphabetical_testing/alphabetical_testing.xlsm) which is used while developing the script, and the other is [Stock Data](./VBA_Stock_data/Multiple_year_stock_data.xlsm) which is the main data to run the script. The data is sourced into Excel Microsoft Office, the test data have 7 sheets(A-P) the size of this file is smaller and used for testing. The stock data or the main data has three sheets categorized yearly 2014, 2015, 2016. This file is bigger in size, it may take time when you execute the script into the file. The VBA scripts also found in both data points file directories.

Create a script that loops through all the stocks for one year and outputs the following information:

The ticker symbol

Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

The total stock volume of the stock. The result should match the following image:

Moderate solution

Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

Hard solution

Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

NOTE
Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

Other Considerations
Use the sheet alphabetical_testing.xlsx while developing your code. This dataset is smaller and will allow you to test faster. Your code should run on this file in under 3 to 5 minutes.

Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.
