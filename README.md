# Module 2 Work & Challenge: Stock Analysis and VBA Code Refactoring
Using VBA to clean up and analyze stock data, along with refactoring code for the quickest run time.

## Overview
### Background
Steve, a finacial advisor to his parents, required assistance in making a repeatable process to analyze stock data from 2017 and 2018 at the click of a button. Using VBA macros, a solution was provided to Steve such that he could press a button, enter the year of stock data he would like to review, and the analysis would be printed up for him with color coded formatting to easily reference how each stock performed that year. After seeing the ease in which Steve could easily analyze stock data, Steve requested the code be refactored to allow for thousands of stocks to be reviewed in seconds. 

### Purpose
Now that a base process of analysis was set up to provide Steve with meaningful information, the next step was to refactor the code and compare run times to make sure the most efficient method of analysis was given to Steve to use on larger datasets. The results below describe the processes of the two solutions, and which one will be most efficient for Steve to use. 


## Methods & Results
### Solution 1
The first solution used nested for loops to first look the rows where the ticker was equal to the ticker beign evaluated, then look at all of the rows of data sum the total volume, find the starting price, and find the ending price. Once those data were found, the data would be inserted into the excel table to display the analytics for Steve to view, and then the next ticker would be evaultated, repeating the process. This process was quite lengthy as the process had to loop through several processes. The run times for the first solution can be seen below. 

![VBA_Chaallenge_Solution1_2017](https://github.com/rmchartman/stock-analysis/blob/master/Resources/VBA_Challenge_Solution1_2017.png) ![VBA_Chaallenge_Solution1_2018](https://github.com/rmchartman/stock-analysis/blob/master/Resources/VBA_Challenge_Solution1_2018.png)

### Solution 2 (Refactored)
The second solution used indexed arrays to use 3 separate for loops that referenced specific tickers. the first loop set the volumes for each ticker to zero. The second loop found the total volume, starting price and ending price that was specific for each ticker by looping through all of the rows of data and adding sata to the arrays for total volumes, starting prices and ending prices. The third loop filled the analysis table in excel with the data from the three arrays that were filled in the previous loop. The run times for the second solution can be seen below. This method proved to be quite fast and should be reccommended for s the final solution for Steve. The run times for the first solution can be seen below.

![VBA_Chaallenge_2017](https://github.com/rmchartman/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png) ![VBA_Chaallenge_2018](https://github.com/rmchartman/stock-analysis/blob/master/Resources/VBA_Challenge_2018.png)

## Summary
By comparing the timed outputs from both solutions, it can be seen that the second solution was much quicker than the first and should be reccommended to Steve for use on larger datasets. By breaking up the loops and using indexes to reference the data to be stored in each array by stock ticker, the computing power required to complete the analysis was significantly less. Due to the factor that the second solution yielded faster results that the first, there are definite advantages to refactoring code to attempt to find different method with the same results. The disadvantages to refactoring code are that it takes additional time, often brings up several errors that may be hard to fix, and may bring a solution that is less robust than the original solution. Applying the pros and cons to this project, it took a significant amount of time to follow the desired instructions and the instructions were difficult to understand, however the results of the refactoring running faster were very satisfying. 
