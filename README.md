# stock-analysis

## Overview
Visual Basic script "yearValueAnalysis.vbs" was created to analyze performance of 12 Green energy stocks in 2017 and 2018. The original script was refactored in order to decrease the time it took to complete the macro with the intention of, eventually, applying this to a much greater number of stocks.

The name of the game is improving readability and decreasing the time it takes for code to run. 

## yearValueAnalysis breakdown
The dataset was presorted with stock tickers in alphabetical order with each stock ordered in increasing chronological order. We first set the stage for the analysis of the stocks by: 
1. initializing several variables like yearValue, startTime, endTime, RowCount, 
2. formatting the worksheet with the desired column headers for output,
3. initializing tickers() array to hold stock tickers.
4. starting the timer

Then yearValueAnalysis script reads and stores total daily volume, starting and ending price into variables by using a nested For loop, first by looping through tickers as iterator i increments. The outer loop calls the ticker string and initializes the totalVolume to zero every time before stepping into the inner For loop.

`insert code with more detailed comment up until the second loop`

The inner For loop uses iterator j to loop through all rows (all 3013 of them!) and interrogates each row with several If statements to determine the values for total volume, starting price and ending price. The For (j) loop concludes and i increments by 1 onto the next ticker. It is important to note here that this method fulfills the If statement by comparing the current column A string with Ticker = Tickers(i).

`insert code with comments`

After the outer loop completes all steps, the data stored in variables Ticker, totalVolume, endingPrice and startingPrice are printed into the output worksheet. The script edits cells in column B for some visual feedback and number formatting. Finally, a msgBox prints the amount of time it took to run the macro. 

(insert png for 2017 and 2018 for original script)

## How did other Green energy stocks fare?
Based on the output of the script, Steve's watchlist of green energy stocks grew in 2017 better than in 2018. This is consistent with events at the time, especially considering the volatile market of 2018 and the steady market decline starting in October 2018. 

## Refactoring 101
In order to optimize the script run-time, the refactored code addressed a few features of the original script.

1. For yearValueAnalysis the nested `For` loop is the biggest bottleneck to a quicker run-time. Specifically, the inner loop of If statements is run `12 tickers * 3013 rows` for a total of 36,156 times. If we adopt the strategy to loop through all rows once only, this will require a different tactic to store data into those output variables (totalVolume, startingPrices and endingPrices). We need a variable that labels or indexes every unique ticker which can also be used in place of an iterator for the various if statements.

Introducing, power duo variable tickerIndex and arrays.
Using variables stores data in memory and is faster to recall. 
2. Does taking out the nested loop make it quicker? I think both sets of code includes 11x3013 for the original vs 1x3013 
3. 

# Why refactor at all? 
## Pros
1. Decimated the runtime
2. Simpler logic-flow, I think? less chance for loop errors and made it more succinct...

## Cons
1. Is the time worth it? Think about the magnitude of the code we are editing, the original code took less than 2 seconds for completion which is hardly a hindrance. While the argument holds true for code much larger with much longer run-times, perhaps in cases like this one it's not worth it.
2. ...but we have to contend with less intuitive logic. Will this be easier to maintain by others? 
3. Does the refactor allow for extension? How will it be able to increase the number of stocks this is applicable to?

# To what extent did refactoring help the original script?
## Pros 

## Cons