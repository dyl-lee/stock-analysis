# Overview
A macro was created to analyze performance of 12 Green energy stocks in 2017 and 2018 "yearValueAnalysis.vbs". The original VBA code was refactored and assessed for efficiency by measuring the time it takes for code to complete the macro. The name of the game is improving readability and decreasing the time it takes for code to run. 


Get pictures of 2017 and 2018 with time and output included (4 pictures)

# yearValueAnalysis breakdown
We first set the stage for the analysis of the tickers by: 
1. initializing several variables like yearValue, startTime and endTime, 
2. formatting the output sheet with the desired column headers
3. initializing tickers() array to hold unique tickers  

# Ways we increased efficiency with refactoring
1. Identify the bottleneck to a quicker run-time. For yearAnalysis 
Using variables
2. Does taking out the nested loop make it quicker? I think both sets of code includes 11x3013 for the original vs 1x3013 
3. 

# Pros
1. Decimated the runtime
2. Simpler logic-flow, I think? less chance for loop errors and made it more succinct...

# Cons
1. Is the time worth it? Think about the magnitude of the code we are editing, the original code took less than 2 seconds for completion which is hardly a hindrance. While the argument holds true for code much larger with much longer run-times, perhaps in cases like this one it's not worth it.
2. ...but we have to contend with less intuitive logic. Will this be easier to maintain by others? 
3. Does the refactor allow for extension?