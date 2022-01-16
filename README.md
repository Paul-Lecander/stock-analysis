# Stock Analysis Using VBA Script

## Overview of Project: Explain the purpose of this analysis.
- In this challenge, the objective was to refactor a VBA script to analyze the return of twelve stocks within two different years. 
- Originally, the macro that was created to look at the twelve stocks but only over one year but now is capable of looking at stocks over two years and at a faster rate. 

	
## Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

- While refactoring the original script, there were many things that were left the same but some slight changes that made a difference. 
An example of this was in the way that the data in the worksheets were looped through.
In the original stock analysis, a *For* loop was created to iterate from 0 to 11 to make looping through the array and accessing each element possible.

![Original_Steps4and5.png](https://github.com/Paul-Lecander/stock-analysis/blob/main/Original_Steps4and5.png) 

- The initial loop goes through each ticker and the nested *For* loop starts looping though all rows in the data. 
Then the initial *If/Then* assessment evaluates if the first cell's value is equivalent to the current ticker, then it is the *Total Volume* + the volume in that row's volume column. 
The second assessment evaluates if the cell's value before the current cell is not equivalent to the current ticker, if it is then the current row in column 6 is the *Starting Price*. 

- Alternatively, the refactored stock analysis uses 4 arrays (one for the 12 tickers and three for outputting data) and uses a variable *tickerIndex* to access all of the arrays. This can be seen in step 1a and 1b in the refactored script.

![Refactor_IndexandArrays.png](https://github.com/Paul-Lecander/stock-analysis/blob/main/Refactor_IndexandArrays.png)
