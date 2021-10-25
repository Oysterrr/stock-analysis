# Stock Analysis Using Excel VBA

## Purpose
The project is to refactor VBA code in order to increase efficiency and measure the performance of 12 different stocks in 2017 and 2018.

## Results
Given the daily trading data of 12 different stocks in 2017 and 2018, an original VBA code is created to output the values of total daily volumes and yearly returns for different stock tickers. A refactored VBA is also created.

### Comparison between 2017 and 2018
According to the two charts created for 2017 and 2018, the increase/decrease of total daily volumes for different stocks between two years are relatively small. While most of the 12 stocks have positive yearly return in 2017, most of the 12 stocks have negative yearly return in 2018 instead.

### Comparison of execution times of the original script and the refactored script
In the original script, the execution time is 0.5546875 seconds to perform the analysis of the year 2017 and 0.5625 seconds of the year 2018.
In the refactored script, the execution time is 0.09375 seconds to perform the analysis of the year 2017 and 0.1015625 seconds of the year 2018.
The speed of running the analysis has increased.

## Summary
Refactoring code has great help improving the efficiency when running the analysis and also organizing the codes. The disadvantage is that refactoring code is time consuming and might lead to a worse solution.