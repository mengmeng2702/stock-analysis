# VBA of Wall Street

## 1. Project Overview

VBA is often used in the financial industry. This project is designed to perform stock data analysis with VBA. Over the past few years, there has been a new concept of green energy. DQ stock, being one of the client's preference, has been analyzed for its overall performance in the year of 2018. With different forms of green energy available in the investment market, this analysis also concludes the result of some other green energy stocks with their total daily volumes purchased and their annual returns. Analysis result will be provided to help the client with their investment choices.

<br/>
With the VBA script created, and refactored, automation was done with code performance measured with elapsed time captured and the screenshots of outcomes attached to show the visual analysis on a variety of stocks for the year 2017 and 2018.

## 2. Result Analysis on Stocks

### All Stocks Analysis - 2017
---
In the year of 2017, all stocks but TERP have a positive yearly return with the top profitable stock being DQ, with the highest annual return (199.4%) among all stocks. The only negative stock of the year is TERP, it has a negative annual return of -7.2%.

<br/>

See below for the 2017 stock returns with positive returns formatted as highlighted in green and negative returns in red:

![VBA_Challenge_2017.png](resources/VBA_Challenge_2017.png)
### All Stocks Analysis - 2018
---
In the year of 2018, most stocks got hit by the economic crisis and have a negative yearly return. Two exceptions are stock ENPH and stock SEDG with yearly return as 81.9% and 84.0% respectively. DQ stock has a -62.6% yearly return on the negative side.

<br/>

See below for the 2018 stock returns:

![VBA_Challenge_2018.png](resources/VBA_Challenge_2018.png)
### All Stocks Analysis - Conclusion
---
Taking both years into consideration, it has shown that ENPH stock not only had done fairly well in the year of 2017 (with a positive return of 129.5%, within top three high returns), it remained as high returns in 2018 even with the attack of the crisis. The stability of the stock return for ENPH has made it the best choice for investors in the upcoming years based on the stock data used in this analysis.

### Code Performance Analysis
---
Based on the original code written, the run time of the code is on average around 0.5s for both 2017 and 2018 all stocks analysis sheet. Please see **below** the screenshots for the original code performance measured with elapsed time for both years.

<br/>

Compare them with the refactored code run time. Please refer to the above stock outcome screenshots taken and the performance measured with elapsed time showing (0.114s for year 2017 and 0.108s for year 2018). There is a comparable difference even for such a small amount of data rows (3013 rows in total).

![VBA_Challenge_2017_original.png](resources/VBA_Challenge_2017_original.png)

<br/>

![VBA_Challenge_2018_original.png](resources/VBA_Challenge_2018_original.png)

## 3. Summary on VBA Scripting

Generally speaking, refactoring code has the advantage of improving the speed of the code running. Whereas original code has the advantage of getting the task done faster.

<br/>

In more details, refactored code has a better, clearer structure which will help the maintaining of it. It also has the potential to be developed more with more complicated code built on top. However, the disadvantages of refactoring code are time consuming, more chances of resulting in errors due to the restructure of the code.