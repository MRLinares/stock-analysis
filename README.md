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

Finally, after a correct year is entered, a message box displays the amount of time it took the program to perform the task.  Below is an example of the time for the original code written before receiving new code that needed refactoring.


#### Original Code
![Original_Code_Run](https://user-images.githubusercontent.com/108758105/183269664-dd166459-f3da-455b-8cb8-f94b05142139.png)

While that is definitely fast, it is based on the program searching data for twelve stocks.  As datasets become increasingly larger, the efficiency at which a program can operate becomes more crucial, so optimizing it at it's simplest form is necessary in order to "future-proof" the code. In the search for program optimization, I came across a site [Eident Training]https://eident.co.uk/2016/03/top-ten-tips-to-speed-up-your-vba-code/ that was instrumental to my successs in optimizing the code (see conclusion).



### Challenges

Refactoring code refers to taking code already written and cleaning it up, improving its functionality and/or performance.  The advantages of this is that there is already a framework put in place for the program to run.  If it is written in a form easy on the eyes, then one could deduce, precisley, the code's itentions and can make their improvements with a lot of information to begin with.  

Unfortunately, this isn't always the case.  According to [BMC Blogs]https://www.bmc.com/blogs/code-refactoring-explained/ sometimes code is written in a rush to wrap up a sprint and release the product; while the code may function as intended, it may have performance issues or may be written in a way that needs cleaning up.  Receiving code like this drastically increases ambiguity and one may have to go through hours of debugging to get the code right.
  
#### Applications to this Project

When dealing with fractions of seconds, it is difficult to truly understand how big of an impact a single line of code can have.  Much of my process involved trial and error through different strategies sourced on Google.  An added challenge is that of refactoring code.  The program, as described above, was not as clean and polished as it is now, in fact, the code ran slower than the intial run of my original code.

#### Refactored Code (Before Improvements)

![1st_Run_2018](https://user-images.githubusercontent.com/108758105/183270054-95bccb7b-f4d8-4045-ae0f-d44ffbe01279.png)

It was challenging finding ways to speed up the process while not only keeping the functionality, but adding user friendly function and simplification.

The most significant boost to performance, I initially thought, was that of the With/End With statements.  Rather than repeating If then statements for the same cell, the "with" statement speeds up the process by allowing the computer to process the commands simultaneously for a given cell or range.


#### Original Code
![Screen Shot 2022-08-08 at 10 53 02 AM](https://user-images.githubusercontent.com/108758105/183447990-bba831c9-8c97-49a6-a53a-462e5fb279ee.png) 


#### Refactored Code (Post Improvements)
![Screen Shot 2022-08-08 at 10 52 17 AM](https://user-images.githubusercontent.com/108758105/183447997-151ca383-f203-42e1-96eb-d2e5135d4bfb.png)
![Screen Shot 2022-08-08 at 10 52 34 AM](https://user-images.githubusercontent.com/108758105/183447993-f7a5a978-edf0-472b-b1bb-30c7ec4ae1e2.png)
![Screen Shot 2022-08-08 at 10 59 39 AM](https://user-images.githubusercontent.com/108758105/183448744-90303c19-676a-45a8-8ba6-e596ecb4b892.png)


### Conclusion

The main factor slowing down my code, however, was the nested loop.  In my original code, I initialized my volume to zero after I started my for loops  (that loop over all the rows) rather than before, essentially creating a nested loop.  By remedying this one piece of code, I drastically improved my code.


Initially, I was able to improve the system performance by roughly 25% and increase it's user functionality via the buttons discussed at the beginning.

#### After Refactoring
![VBA_Challenge_2017](https://user-images.githubusercontent.com/108758105/183270004-4995e92e-adab-4642-b181-13bc5a495b10.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/108758105/183270009-6ee2a035-d7a3-4078-9d02-29c5e1d47bfa.png)

After the upgrade, though I was able to achieve over 90% improvement!

![VBA_Challenge_2017_2](https://user-images.githubusercontent.com/108758105/183800281-bea5a0fb-9f40-4328-9a87-15b0c4ca7703.png)  ![VBA_Challenge_2018_2](https://user-images.githubusercontent.com/108758105/183800299-a7497720-d0ed-4d1b-8f98-aabed92e2809.png)


