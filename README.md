# StockData_VBA
* In this repo I used VBA code (Excel Macros) to automate processes in excel. 
* The joy of VBA is to take the tediousness out of repetitive tasks and run them over and over again with a click of the button. That is exactly what the VBA code in this repo does. 
* The VBA code in this Macro creates annual summary tables for stock data. The code finds the yearly change, yearly percentage change, and the total stock volume traded for every single stock in the table with the click of a button rather than having to do this manually. Additionally when you click the button to execute the macro, the macro creates a second summary table with the stock with the greatest percentage increase that year, the stock with the greatest percentage decrease that year, and the stock with the greatest total volume traded that year. 

## Starting Data 
* I started this repo with 2 excel spreadsheets, "Alphabetical_Testing" and "Multiple_Year_Stock_Data", both of which can be found in the folder "01 Before Running Macros". "Alphabetical_Testing.xlsm" is a sample of "Mulitple_Year_Stock_Data.xlsm", it only contains data for 2016 whereas "Multiple_Year_Stock_Data.xlsm" contains data for 2014, 2015, and 2016. I chose to make this sample of the dataset because macros can be slow when working with a lot of data. So to save time I used the sample dataset, Alphabetical_Testing, to test my macro code and make sure that it worked sucessfully before running the macro code on the entire dataset, Mulipte_Year_Stock_Data.

*  Both spreadsheets contain the same columns ticker, date, open, high, low, close, and vol. Here is a screenshot of some rows from the dataset. 

![](04_README_IMAGES/Sample_Starting_Data.JPG)
