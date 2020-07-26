# Module 2 Work & Challenge: Stock Analysis and VBA Code Refactoring
Using VBA to clean up and analyze stock data, along with refactoring code for the quickest run time.

## Overview
### Background
Steve, a finacial advisor to his parents, required assistance in making a repeatable process to analyze stock data from 2017 and 2018 at the click of a button. Using VBA macros, a solution was provided to Steve such that he could press a button, enter the year of stock data he would like to review, and the analysis would be printed up for him with color coded formatting to easily reference how each stock performed that year. After seeing the ease in which Steve could easily analyze stock data, Steve requested the code be refactored to allow for thousands of stocks to be reviewed in seconds. 

### Purpose
Now that a base process of analysis was set up to provide Steve with meaningful information, the next step was to refactor the code and compare run times to make sure the most efficient method of analysis was given to Steve to use on larger datasets. The results below describe the processes of the two solutions, and which one will be most efficient for Steve to use. 


## Results
### Solution 1
The first solution used nested for loops to first look the rows where the ticker was equal to the ticker beign evaluated, then look at all of the rows of data sum the total volume, find the starting price, and find the ending price. Once those data were found, the data would be inserted into the excel table to display the analytics for Steve to view, and then the next ticker would be evaultated, repeating the process. This process was quite lengthy as the process had to loop through several processes. The run times for the first solution can be seen below. 
![VBA_Chaallenge_Solution1_2017](https://github.com/rmchartman/stock-analysis/blob/master/Resources/VBA_Challenge_Solution1_2017.png)

![VBA_Chaallenge_Solution1_2018](https://github.com/rmchartman/stock-analysis/blob/master/Resources/VBA_Challenge_Solution1_2018.png)


![VBA_Chaallenge_2017](https://github.com/rmchartman/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png)

![VBA_Chaallenge_2018](https://github.com/rmchartman/stock-analysis/blob/master/Resources/VBA_Challenge_2018.png)

## Summary
Answer the following questions:
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
