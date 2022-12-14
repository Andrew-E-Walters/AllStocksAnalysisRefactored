# AllStocksAnalysisRefactored
## Refactor VBA code and measure performance
### Purpose
Steve wants to expand the dataset to include the entire stock market over the last few years. While he likes our inital code, he wants to refactor it so that it can handle more than 12 Stocks if needed in the future. We are tasked with creating a code that would work for this data set, and also handle a much larger data set for future use. 
## Overview of Project 
The analysis was conducted by pulling together data from the stocks tab. We were tasked to refactor the Module2_VBA_Script so you loop through the data one time and collect all of the information. The goal is to see if this information will be pulled faster with the new code we are creating. The logic is that if it is faster with this data, it would be better to use the new code with larger data sets going forward, as well as condense our code. 
## Analysis and Results
### Analysis of Code
Initally the code looks similar to what we had created through the module. Though we will have to make some adjustments. We can see from the begining that the code is incomplete and we will have to fill in some sections to get it to work.  
#### Below is the initial code
![Starting Code](https://github.com/Andrew-E-Walters/AllStocksAnalysisRefactored/blob/main/Original_Code.png) 
### Changes of Inital Code
We had to create a tickerIndex variable and set it equal to zero. We then can use the tickerIndex to access the correct index across the different arrays. The three output arrays tickerVolumes, tickerStartingPrices, and tickerEndingPrices had to be created. We then needed to ensure that they were the right data types. After that we had to create for loops to set the tickerVolumes to zero, and another for loop to loop over all rows in the spreadsheet. We then had to increase the current tickerVolumes by using the tickerIndex variable as the index. The next step was to create a series of if-then statments to get the startingPrices and endingPrices. Once we have that infromation we are good to increase the tickerIndex. Finally we use a for loop to loop through our arrays and output the Ticker, Total Daily Volume, and Return columns in the spreadsheet we are using to view the results. 

![Final Code](https://github.com/Andrew-E-Walters/AllStocksAnalysisRefactored/blob/main/New_Code.png)

There were some results that came with this updated code as well as some Advantages and Disadvantages. 

### Results 
When we ran the All Stocks Analysis for the module, the texbox displayed that the code ran for 0.60015625 seconds for the year 2017. 

![2017 Original](https://github.com/Andrew-E-Walters/AllStocksAnalysisRefactored/blob/main/All_Stocks_Analysis_2017.png)

When we ran the code we refactored the code ran for 0.2617188 seconds in the year 2017, a whole 0.33843745 seconds faster!

![2017 Refactored](https://github.com/Andrew-E-Walters/AllStocksAnalysisRefactored/blob/main/VBA_Challenge_2017.png)

When we ran the All Stocks Analysis for the module, the texbox displayed that the code ran for 0.703125 seconds for the year 2018.

![2018 Original](https://github.com/Andrew-E-Walters/AllStocksAnalysisRefactored/blob/main/All_Stocks_Analysis_2018.png)

When we ran the code we refactored the code ran for 0.1171875 seconds in the year 2018, a whole 0.5859375 seconds faster!

![2017 Refactored](https://github.com/Andrew-E-Walters/AllStocksAnalysisRefactored/blob/main/VBA_Challenge_2018.png)


It is clear that the code we created is faster than the previous one. Not only that but we have also made some adjusments in formatting that allows us to clearly display the data in a way that is visually appealing. 

### Stocks Analysis
On top of making the code faster we also created a code that displays this information in a simple way that someone not familiar with the topic could easily interprit. The addition of the conditional formatting to display the performance of the stock allows us to see what stocks had a good vs a bad retun for their respective years.
#### The 2017 Performance had most stocks increasing their value over the course of the year. 
![2017 Stocks](https://github.com/Andrew-E-Walters/AllStocksAnalysisRefactored/blob/main/2017%20Stock%20Performance.png)
#### The 2018 Performance had most stocks decreasing their value over the course of the year. 
![2018 Stocks](https://github.com/Andrew-E-Walters/AllStocksAnalysisRefactored/blob/main/2018%20Stock%20Performance.png)

This information could be sent to Steve and he would be able to tell what stocks he might want to invest in, and what he might want to stay away from in the upcoming years based on the 2017 and 2018 performance. 
## Summary
### Advantages
As Stated earlier, the code that we created by refactoring the code was faster than what was created in the module. This would allow us to run this code much faster over a larger set of data. This code was only looping over 2 to 3013 but if it was more than that it would take more time with the inital code. This code also has formatted text depending on the return of the stock which allows for a clearer visualization of the data. 
### Disadvantages 
Being inexpirenced in Coding, this took a good amount of time to format and debug. Initally I could only get the first value to display, then I managed to get it to run but it would stop at the 12th line of rows. Then I had to change certian variables so that the data would not be gathering inccorect values. Initially I had everything set to (i) which was causing issues. After much trial and error I was able to get the code to run and display the correct information, but this was after some time fixing the code to display the data. It would have initally been faster to use my old code rather than wait for me to create a code that was a fraction of a second faster. In the long run, the new codes creation will be more useful for larger applications. 

