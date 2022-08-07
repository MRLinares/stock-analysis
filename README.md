# Stock Analysis
A client is performing brokerage services and needs to see how several stocks have performed for two years. Visual Basic in Excel was used for the financial analysis of the stock data.  Code was originally written to perfom the analysis, but was refactored to improve efficiency.
### Process

Data for financial stocks were downloaded and filtered for twelve select stocks.  The twelve stocks were compiled into an array accessible through a ticker index.  The code was written so that the program would read through all the stocks and sum up all the daily volumes for each requested stock and display total daily volume and return percentages.  Code was also written to format a table to provide information that is easily accessible for the client

![Example_Empty_Table](https://user-images.githubusercontent.com/108758105/183268838-0af21f2e-4684-4eb3-8e2e-c42ef62bdc28.png)

Buttons were programmed so that the client would never have to open a developer tab in Microsoft Excel and operate macros.  The buttons are designed to run an analysis for a desired year, highlight the returns so that negative returns are red and positive returns are green, remove the highlights for optical preferences, and to reset the entire table, respectively.

When pressing the "Run Analysis" button, an input box will appear requesting the year to analyze

![Example_Inputbox](https://user-images.githubusercontent.com/108758105/183269386-50d18bc6-de5a-4e6f-9552-9cbfac1a83b0.png)

Entering a valid year will populate the table with the total daily volume and return percentages. For example, below is a sample for the year 2017:

![Example_2017_Run](https://user-images.githubusercontent.com/108758105/183269458-b1996f56-51e1-4644-88b8-ea8aa928e914.png)

Entering a year not within the dataset returns an error message

![Example_error](https://user-images.githubusercontent.com/108758105/183269430-734e04a1-daa2-4444-afb9-e1f30e04ef89.png)




### Challenges
When dealing with fractions of seconds, it is difficult to truly understand how big of an impact a single line of code can have.  Much of my process was trial and error trhough different strategies sourced on Google.
