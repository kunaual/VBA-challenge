This folder contains 4 files:
	This file- ReadMe.md
	StockCalcVBA.bas
	StockData_xlsm_screenshots.jpg
	Multiple_year_stock_data.xlsx
	
StockCalcVBA.bas-
	I wrote this VBA script to calculate 3 values for each ticker for each year (sheet):
		* Yearly change (difference between ticker's year closing price and year open price)
		* Percent change of stock price
		* Total stock volume
		
	For each year, the script highlights the ticker and value with the biggest increase (%) in price, biggest decrease (%), and highest total volume in columns O and P. 
	
	The script assumes that stock ticker data on each sheet are grouped together by ticker and ordered by date.  i.e. all ticker "A" rows are next to each other Jan through Dec, all "AA" rows are together and so on. 

StockData_xlsm_screenshots.jpg-
	Screenshots of Multiple_year_stock_data.xlsx with the StockCalcVBA script run against it.


Multiple_year_stock_data.xlsx-
	Spreadsheet of stock data 2014-2016 for the script to run against.