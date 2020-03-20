# StockAnalysis_Module_2

## Introduction
The code presented is to analyze stock data and summarize the performance of each stock. The data provided is daily stock price for years 2017and 2018. The user's requirement is a VBA macro, that are accessed with easy to use buttons in the worksheet, that lets them select a year and generate a table of the volume and gain/loss over the year for each stock. An additinal requirement is to format the output table with red/green colors so the user can interpret performance data quickly, at a glance.

## Need for refactoring
The initial approach was to create a single array of tickers and parse the full data separately for each ticker to extract starting price, closing price, and volume. This approach resulted in parsing the large data set 12 time since there are 12 stocks. This process is repetiitve and computationally expensive. I refactored the code to improve the performance and make the process more efficient.

## Refactoring approach
The approach I followed is comprised of three steps:
- Create an array for each data column required 
- Extract all the data needed from each row with one parse of that row
- Analyze data, provide user interaction buttons, and apply formats as previously so th euser experience is the same.

## Code flow
(1) User provides input year for analysis: This lets user specify which year they would like to analyze  
(2) Add header and formats: Header text and column formats are set  
(3) Activate sheet and define variables/arrays: The sheet corresponding to user's choice is selected. I defined six array variables and an integer that are used in the code.  
(4) Initialize arrays: I used two types of inititalizations. (a) The first type is to provide an ordered list of contents of the tickers() array; (b) I used a 'for' loop to initialize values for the other arrays.  
(5) Determine last row: This part determines the dimensions of the data set by determining the last row (so it does not need to be specified in the code by number).  
(6) Populate arrays: For each array (other than tickers()), values are updated from initial based on conditions by parsing each row in the data set only once.  
(7) Write and format results: I created a separate worksheet 'All Stocks Analysis_refactored' to write results from each applicable array. I also formatted the results conditionally as per the requirement.


