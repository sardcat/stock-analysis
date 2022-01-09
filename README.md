# MODULE CHALLENGE 2 : VBA Stock Analysis

## Introduction & Purpose

Friendly Steve is interested in helping his parents with choosing stocks in the Green Energy sector. Having provided "The Team" with an
Excel worksheet of recent year stock data for the stocks he's chosen to research, data analysis utilizing VBA to quickly highlight for Steve
stock performance is rendered. The code was also written with flexibility should he choose to expand his dataset. This Challenge focuses on 
refactoring code to improve performance via ARRAYS over nested FOR loops.

## Results

By making use of three Arrays:```tickerVolumes(), tickerStartingPrices(), tickerEndingPrices()``` a nested FOR loop was able to be removed. This
resulted in the program only having to pass through the Worksheet data once.

The reduction in processing time was the following:
![2017](/VBA_Challenge_2017.png)
![2018](/VBA_Challenge_2018.png)

By comparison using a nested FOR loop, the 2017 and 2018 processing times were 1.175781 seconds apiece. This is nearly a full second removed in each iteration.

## Summary

Refactoring code can have the benefits of making the code more legible when viewed by additional sources or streamline processing power, especially
as additional computing features become available. However, there is a significant time investment by the coder(s) to research new avenues of approach and
then implement the changes, especially if the difference in processing time is minimal or the code won't be heavily utilized.

In this Module Challenge, removing the nested FOR loop and replacement with three ARRAYS noticeably improved time to process. While ARRAYS are powerful in making
use of computer memory, they could in theory quickly expand in size if not properly contained and wastefully occupy computing resources, providing no additional
performance increases over straightforward single variable use in nested FORs.
