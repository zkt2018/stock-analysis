#**stock-analysis**
##**Overview of Project**
Through this module, we used VBA to analyse a stocks dataset to provide the required information for Steve. Steve needs to evaluate some stock data to advise his parents for their investments. For this project, we evaluated a specific company’s volumes and returns, DAQO (with the ticker symbol: *DQ*) following with other companies to find out which had positive returns in 2017 and 2018.
This dataset includes the opening and closing prices, and daily volumes for 12 tickers. Using VBA, we ran an analysis through all tickers to get their Total Daily Volumes and Returns.
##**Results**
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
For the project, we used two different codes to analyse the data. First we calculated the daily volumes and return for the *DQ* company as this one is specified by Steve’s parents. However, *DQ*’s return in 2018 has been negative. Hence, we decided to evaluate other companies' returns in 2017 and 2018 as well.
Based on the results, all the companies had positive returns in 2017 except for *TERP*, while their return has significantly dropped in 2018 except for *ENPH* and *RUN*.
[VBA_Challenge](https://github.com/zkt2018/stock-analysis/blob/main/VBA_Challenge.xlsm)
As we often are dealing with huge datasets, the run time of the code is considered as an important factor when using Macros. We tried two methods; one included nested for loops and had multiple subroutines versus the other one where we included two separate for loops. As displayed in the following images, the run times were remarkably different with the two methods. The first code was running within about 0.79 seconds with the first code, whereas with the second method, it took about 0.10 seconds.
First code running times (2017 and 2018)

![Initial_Code_Running_Time_2017](https://github.com/zkt2018/stock-analysis/blob/main/Resources/Initial_Code_Running_Time_2017.PNG)
![Initial_Code_Running_Time_2018](https://github.com/zkt2018/stock-analysis/blob/main/Resources/Initial_Code_Running_Time_2018.PNG)

Vs. second code running times (2017 and 2018)

![VBA_Challenge_2017](https://github.com/zkt2018/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](https://github.com/zkt2018/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

However, the run times slightly fluctuate using different methods to run the code, i.e. that can be often less when we run the code using the *play* button in VBA instead of clicking on the macro button created on the worksheet. It also differs when removing the *for* loop where we have initialized the totalVolumes as below example:
![VBA_Challenge_2017_WithoutForLoop](https://github.com/zkt2018/stock-analysis/blob/main/Resources/VBA_Challenge_2017_WithoutForLoop.png)

Vs. with *For* loop:
![VBA_Challenge_2017_forLoop](https://github.com/zkt2018/stock-analysis/blob/main/Resources/VBA_Challenge_2017_forLoop.png)

Since the run time slightly fluctuates between 0.27 and 0.78 seconds by adding/removing and by restarting the workbook or deleting the content and re-running, I let the *for* loop be in the code.

##**Summary**
###What are the advantages or disadvantages of refactoring code?
By refactoring the code, we got the same results in a more efficient manner and within a less time. This is important specially when we are analysing huge datasets and where we need to save time. 
A disadvantage of refactoring, however, can be ending up with a longer piece of code in one subroutine as we have tried to use less nested loops. 
###How do these pros and cons apply to refactoring the original VBA script?
In this module, through refactoring, we loop through all the data in order versus loop over some part of data in another loop. The nested loop multiplies the number of times the processor needs to go through data to check for the conditions which is why the run time is longer than when it loops through each cell only once.
Although the first code might seem more understandable, the second one is more efficient.
