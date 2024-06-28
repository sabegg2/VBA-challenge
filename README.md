# VBA-challenge
Module 2 Challenge

I have included:
(1) A text file of my macro, titled Sub QuartlyStockAnalysis().rtf
(2) Four screenshots of the Multiple_year_stock_data.xlsm file showing that I achieved the desired output. Each screenshot corresponds to a different quarter (Q1, Q2, Q3, Q4).
(3) The complete excel file with my macro embedded, titled Multiple_year_stock_data_edited.xlsm.

I wrote the code myself, frequently using Google and Microsoft Excel help to determine how to write various methods to do what I wanted. I did not get any help from any online AI programs. The main piece of coding I did not write myself was the line of code for finding the last row of data, which we saw in class on June 27, and is also found here: https://techcommunity.microsoft.com/t5/excel/understand-cells-rows-count-a-end-xlup-row/m-p/292728

Note that my code is based on the spreadsheet being sorted as it is given, i.e., by Ticker and by Date. If the spreadsheet were not sorted this way, the the spreadsheet should first be sorted by Ticker and then by Date. Sorting the spreadsheet this way will be much quicker than executing a for loop that has to loop through all of the data for each ticker and find the minimum and maximum dates and respective opening and closing prices for each ticker.

A couple of concerns with the rubric:
(1) The rubric mentions conditional formatting for the percent change column, but this was not requested nor shown in the example solution outputs. Only conditional formatting for the quarterly change was requested and shown.
(2) The rubric mentions that the script loops through the stock data and reads/stores the ticker symbol, volume, open price, and closing price from each row. However, since the data is sorted by increasing date, it is only necessary to get the open price from the first row for a ticker and the closing price from the last row for a ticker, and grab the ticker only once as well. This is what I have done in my code.
