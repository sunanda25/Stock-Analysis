## Overview
Steve’s parents are very passionate about Green energy and they are looking forward to investing money in it. Without conducting much research, they decide to invest all their money in “DAQO New Energy Corp”. Steve is concerned about their parents not diversifying their stocks, so he decided to build an Excel sheet for 2017 and 2018 stocks of 12 different Green energy companies. To help Steve analyze this data, an efficient tool more complex than Excel formulas is required. VBA which is an extension to Excel has been used for this purpose. An Excel macro has already been provided to do the stock analysis; however, the code is neither optimized nor easily readable. The original macro has been refactored by adding new variables and arrays which resulted in more readability and faster performance of the macro. 

## Results
### Refactoring the code
To improve efficiency, the original code has been refactored by switching the order of the loops. 

1.	Created a tickerIndex variable as a Single data type and assigned a value of Zero.


2.	Three output arrays are created:
   *	The tickerVolume array as a Long data type
   *	The tickerStartingPrices array as a Single data type
   *	The tickerEndingPrices array as a Single data type



3.	Created a For loop to initialize the tickerVolumes to Zero.


4.	Created a For loop over all the rows in a spreadsheet by increasing the tickerVolumes using the tickerIndex. Using If-Then statements:
   *	If the current row is the first row of the selected tickerIndex, then assign the current closing price to the tickerStartingPrices.
   *	If the current row is the last row of the selected tickerIndex, then assign the current closing price to the tickerEndingPrices.
   *	Increase the tickerIndex if the next row ticker does not match the previous row ticker.
   
   
5.	Created a For loop using tickerIndex to display the outputs as Ticker, Total daily volume, Return.
 
 
### Stock Performance of 2017 and 2018
*	Stock returns for 2017 are positive except for “TERP”.
*	Stock returns for 2018 are mostly negative except for “ENPH”, “RUN”.
*	DQ ticker has a 199.4% return in 2017 whereas in 2018 it has -62.6%.



### Execution Times
The execution time of original code for the years 2017 and 2018:


![tickerIndex](https://user-images.githubusercontent.com/76491891/110210081-06244d00-7e5e-11eb-9300-3df8088c8622.png)

The execution time of refactored code for the years 2017 and 2018:


## Summary
### Advantages or Disadvantages of Refactoring code:
*	The refactored code is well structured and easy to read. 
*	The refactored code works more efficiently than the original code when working with large sets of data.
*	The run time for the refactored code is faster than the original code run time.
*	A disadvantage of refactoring code is new bugs can be introduced.

### Pros and cons for refactoring the original VBA script:
*	Refactoring the original code makes the macro run faster. The run times of the original code in 2017 and 2018 are 0.731562500004657, 0.758749999993597. After refactoring the code, run times of 2017 and 2018 are 0.1171875, 0.1328125 which are much faster (~600%) compared to original code run times.
*	The refactored code is more readable with appropriate variable names and resequencing of ‘For’ loops.
