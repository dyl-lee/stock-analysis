# Overview
Visual Basic script "yearValueAnalysis.vbs" was created to analyze performance of 12 Green energy stocks in 2017 and 2018. The original script was refactored in order to decrease the time it took to complete the macro with the intention of, eventually, applying this to a much greater number of stocks.

The name of the game is improving readability and decreasing the time it takes for code to run. 


Get pictures of:
* 2017 and 2018 with time and output included (4 pictures)
* yearValueAnalysis code? (not pic)

# yearValueAnalysis breakdown
The dataset was presorted with stock tickers in alphabetical order with each stock ordered in increasing chronological order. We first set the stage for the analysis of the stocks by: 
1. initializing several variables like yearValue, startTime, endTime, RowCount, 
2. formatting the worksheet with the desired column headers for output,
3. initializing tickers() array to hold stock tickers.

Then yearValueAnalysis script reads and stores total daily volume, starting and ending price into variables by using a nested For loop, first by looping through tickers as iterator i increments. The outer loop calls the ticker string and initializes the totalVolume to zero every time before stepping into the inner For loop.

(insert code with more detailed comment up until the second loop)

The inner For loop uses iterator j to loop through all rows (all 3013 of them!) and interrogates each row with several If statements to determine the values for total volume, starting price and ending price. The For (j) loop concludes and i increments by 1 onto the next ticker. It is important to note here that this method fulfills the If statement by comparing the current column 1 string with Ticker = Tickers(i).

(insert code with comments)

After, the outer loop finishes all steps, the data stored in variables Ticker, totalVolume, endingPrice and startingPrice are printed into the output worksheet. 


# Ways we increased efficiency with refactoring
1. Identify the bottleneck to a quicker run-time. For yearValueAnalysis this is the looping through all rows for each ticker. 
Using variables stores data in memory and is faster to recall. 
2. Does taking out the nested loop make it quicker? I think both sets of code includes 11x3013 for the original vs 1x3013 
3. 

# Pros
1. Decimated the runtime
2. Simpler logic-flow, I think? less chance for loop errors and made it more succinct...

# Cons
1. Is the time worth it? Think about the magnitude of the code we are editing, the original code took less than 2 seconds for completion which is hardly a hindrance. While the argument holds true for code much larger with much longer run-times, perhaps in cases like this one it's not worth it.
2. ...but we have to contend with less intuitive logic. Will this be easier to maintain by others? 
3. Does the refactor allow for extension?