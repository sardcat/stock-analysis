# MODULE CHALLENGE 2 : VBA Stock Analysis

## Introduction & Purpose

Friendly Steve is interested in helping his parents with choosing stocks in the Green Energy sector. Having proveded "The Team" with an
Excel worksheet of recent year stock data for the stocks he's chosen to research, data analysis utilizing VBA to quickly highlight for Steve
stock performance is rendered. The code was also written with flexibility should he choose to expand his dataset. This Challenge focuses on 
refactoring code to improve performance via ARRAYS over nested FOR loops.

## Results

By making use of three Arrays:```tickerVolumes(), tickerStartingPrices(), tickerEndingPrices()``` a nested FOR loop was able to be removed. This
resulted in the program only having to pass through the Worksheet data once. The reduction in processing time was the following:
![2017](Resources/VBA_Challenge_2017.png)
![2018](Resources/VBA_Challenge_2018.png)


