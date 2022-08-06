# AllStocksAnalysisRefactored
## Refactor VBA code and measure performance
### Purpose
Steve wants to expand the dataset to include the entire stock market over the last few years. While he likes our inital code, he wants to refactor it so that it can handle more than 12 Stocks if needed in the future. 
## Overview of Project 
The analysis was conducted by pulling together data from the stocks tab. We were tasked to refactor the Module2_VBA_Script so you loop through the data one time and collect all of the information. The goal is to see if this information will be pulled faster with the new code we are creating. 
## Analysis and Results
### Analysis of Code
Initally the code looks similar to what we had created through the module. Though we will have to make some adjustments. We can see from the begining that the code is incomplete and we will have to fill in some sections to get it to work.  
#### Below is the initial code
![Graph of Outcomes](https://github.com/Andrew-E-Walters/Kickstarting-With-Excel-Project/blob/main/Theater_Outcomes_vs_Launch.png) 
### Changes of Inital Code
We had to create a tickerIndex variable and set it equal to zero. We then can use the tickerIndex to access the correct index across the different arrays. THe three output arrays tickerVolumes, tickerStartingPrices, and tickerEndingPrices had to be created. We then needed to ensure that they were the right data types. After that we had to create for loops to set the tickerVolumes to zero, and another for loop to loop over all rows in the spreadsheet. We then had to increase the current tickerVolumes by using the tickerIndex variable as the index. The next step was to create a series of if-then statments to get the startingPrices and endingPrices. Once we have that infromation we are good to increase the tickerIndex. Finally we use a for loop to loop through our arrays and output the Ticker, Total Daily Volume, and Return columns in the spreadsheet we are using to view the results. 
### Results 
When we ran the All Stocks Analysis for the module, the texbox displayed that the code ran for 0.60015625 seconds for the year 2017. When we ran the code we refactored the code ran for 0.2617188 seconds in the year 2017. It is clear that the code we created is faster than the previous one. 
## Summary
### Advantages
mkes coe faster, we were able to get information faster and we can see the values displayed with the 
### Disadvantages 
Hard to fix, takes time. in this case only slihtly faster 

