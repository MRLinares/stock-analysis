# Stock Analysis
A client is performing brokerage services and needs to see how several stocks have performed for two years. Visual Basic in Excel was used for the financial analysis of the stock data.  Code was originally written to perfom the analysis, but was refactored to improve efficiency.
### Process

Data for financial stocks were downloaded and filtered for twelve select stocks.  The twelve stocks were compiled into an array accessible through a ticker index.  The code was written so that the program would read through all the stocks and sum up all the daily volumes for each requested stock and display total daily volume and return percentages.  Code was also written to format a table to provide information that is easily accessible for the client

![Example_Empty_Table](https://user-images.githubusercontent.com/108758105/183268838-0af21f2e-4684-4eb3-8e2e-c42ef62bdc28.png)

Buttons were programmed so that the client would never have to open a developer tab in Microsoft Excel and operate macros, though a macro-enabled file would have to opened to operate the functions.  The buttons are designed to run an analysis for a desired year, highlight the returns so that negative returns are red and positive returns are green, remove the highlights for optical preferences, and to reset the entire table, respectively.

When pressing the "Run Analysis" button, an input box will appear requesting the year to analyze

![Example_Inputbox](https://user-images.githubusercontent.com/108758105/183269386-50d18bc6-de5a-4e6f-9552-9cbfac1a83b0.png)

Entering a valid year will populate the table with the total daily volume and return percentages. For example, below is a sample for the year 2017:

![Example_2017_Run](https://user-images.githubusercontent.com/108758105/183269458-b1996f56-51e1-4644-88b8-ea8aa928e914.png)

Entering a year not within the dataset returns an error message

![Example_error](https://user-images.githubusercontent.com/108758105/183269430-734e04a1-daa2-4444-afb9-e1f30e04ef89.png)

Finally, after a correct year is entered, a message box displays the amount of time it took the program to perform the task.  Below is an example of the time for the original code written before refactoring.

![Original_Code_Run](https://user-images.githubusercontent.com/108758105/183269664-dd166459-f3da-455b-8cb8-f94b05142139.png)

While that is definitely fast, it is based on the program searching data for twelve stocks.  As datasets become increasingly larger, the efficiency at which a program can operate becomes more crucial, so optimizing it at it's simplest form is necessary in order to "future-proof" the code. In the search for program optimization, I came across a site [Eident Training]https://eident.co.uk/2016/03/top-ten-tips-to-speed-up-your-vba-code/ that was instrumental in my successs and I was able to improve significant time.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/108758105/183270004-4995e92e-adab-4642-b181-13bc5a495b10.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/108758105/183270009-6ee2a035-d7a3-4078-9d02-29c5e1d47bfa.png)


### Challenges
When dealing with fractions of seconds, it is difficult to truly understand how big of an impact a single line of code can have.  Much of my process was trial and error through different strategies sourced on Google.  An added challenge is that of refactoring already written code.  The program, as described above, was not as clean and polished as it is now, in fact, the code ran slower than the intial run of my original code.

![1st_Run_2018](https://user-images.githubusercontent.com/108758105/183270054-95bccb7b-f4d8-4045-ae0f-d44ffbe01279.png)

It was challenging finding ways to speed up the process while not only keeping the functionality, but adding user friendly function and simplification.
