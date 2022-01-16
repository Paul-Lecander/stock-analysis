# Stock Analysis Using VBA Script

## Overview of Project
- In this challenge, the objective was to refactor a VBA script to analyze the return of twelve stocks within two different years. 
- Originally, the macro that was created to look at the twelve stocks but only over one year but now is capable of looking at stocks over two years and at a faster rate. 

	
## Results

### Comparing Original and Refactored Sccripts

- While refactoring the original script, there were many things that were left the same but some slight changes that made a difference. 
An example of this was in the way that the data in the worksheets were looped through.
In the original stock analysis, a *For* loop was created to iterate from 0 to 11 to make looping through the array and accessing each element possible.

![Original_Steps4and5.png](https://github.com/Paul-Lecander/stock-analysis/blob/main/Original_Steps4and5.png) 

- The initial loop goes through each ticker and the nested *For* loop starts looping though all rows in the data. 
Then the initial *If/Then* assessment evaluates if the first cell's value is equivalent to the current ticker, then it is the *Total Volume* + the volume in that row's volume column. 
The second assessment evaluates if the cell's value before the current cell is not equivalent to the current ticker, if it is then the current row in column 6 is the *Starting Price*. 

- Alternatively, the refactored stock analysis uses 4 arrays (one for the 12 tickers and three for outputting data) and uses a variable *tickerIndex* to access all of the arrays. This can be seen in step 1a and 1b in the refactored script.

![Refactor_IndexandArrays.png](https://github.com/Paul-Lecander/stock-analysis/blob/main/Refactor_IndexandArrays.png)

- Another key difference was the use of two separate *For* loops. The first *For* loop (2a) was to initialize the *tickerVolumes* to 0 after each loop. The second *For* loop that was created was to loop over all rows in the spreadsheet.

![Refactor_Steps2to3.png](https://github.com/Paul-Lecander/stock-analysis/blob/main/Refactor_Steps2to3.png)

- The steps 3a through 3c were similar to that in the original stock analysis except the *ticker* was being iterated by the *tickerIndex*. At the end of the loop (3d) the *tickerIndex* increased to the next *ticker* after the last cell's value was not equal to the current *ticker*.

### Assessing Original and Refactored Script Speeds
 	  
- The original script took 0.863 seconds to run through all of the data and output it into a new worksheet.

![VBA_Challenge_Originaltime_2018.png](https://github.com/Paul-Lecander/stock-analysis/blob/main/VBA_Challenge_Originaltime_2018.png)

- The refactored scripts were much faster at completing the task with the 2017 analysis taking only 0.125 seconds and the 2018 analysis taking 0.363 seconds. 

![VBA_Challenge_2017.png](https://github.com/Paul-Lecander/stock-analysis/blob/main/VBA_Challenge_2017.png)

![VBA_Challenge_2018.png](https://github.com/Paul-Lecander/stock-analysis/blob/main/VBA_Challenge_2018.png)
