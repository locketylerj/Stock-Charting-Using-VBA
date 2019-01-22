

![stock Market](Images/stockmarket.jpg) 
### This project is a continuation of the Ty_VBA_MultipleYear_StockData-Analysis project. It includes dummy data for multiple made-up stocks for the year 2016. 

### The HStockAnalysis Module includes the maximum and minimum values for percent changes and the maximum total volume along with all the unique stock ticker symbols. 

* This script will loop through all the stocks and take the following info:

  * Yearly change from what the stock opened the year at to what the closing price was.

  * The percent change from the what it opened the year at to what it closed.

  * The total Volume of the stock

  * Ticker symbol

* It is also conditionally formatted to highlight positive change in green and negative change in red.

* This module will also  locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume" and puts them in a separate area on each worksheet starting in cell "O2".

### The ClearResults Module simply erases the data summarized in the HStockAnalysis Module.

### The FilterCopy Module filters all the data in the "Data" worksheet based on the ticker symbol selected in the dropdown. 

* The filtered data is then copied to the OHLCChart tab so that a chart can be made for that particular stock data.

* The filter is then removed to show the data for all stocks on the Data tab. 

### The CreateOHLC Module creates a VOHLC chart for the data copied to the OHLCChart tab. 

* Any data copied to this tab will be charted when the "Create VOHLC Chart" button is clicked.

* The Chart is formatted with Price on the right vertical axis and Volume on the left vertical axis. 

* The date range that is charted on the horizontal axis is only the first 6 months of the year 2016 from January 1st through June 30th. This makes the chart visual more viewable and easily understood. 

* The next steps will be to include date range drop downs that will allow the user to select what date range they would like to chart so that one month to the entire year can be charted if desired. 









