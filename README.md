# Kickstarting with Excel

## Overview of Project

    Our original project was to design a stock analysis tool in Excel, using VBA, that would take a spreadsheet of daily stock figures for twelve stocks and calculate the annual volume and annual return for each stock.  In addition, the results would be formatted to separate stocks based on the annual return, providing an easy visualization for the client to make financial recommendations on.

    The purpose of this project is to scale up the capability of the original project to produce the same results for a greater number of stocks that the client may want to analyze.  In scaling up the potential number of inputs to use in this analysis, we are tasked with refactoring the original code to ensure that is runs more efficiently than the original.  

## Results

    
    
    
    
    Results: Using images and examples of your code, compare the stock performance between 2018 and 2018, as well as the execution times of the original script and the refactored script.

## Summary

### Advantages and Disadvantages of Refactoring Code
    Refactoring is an important step in any code design of an algorithm.  The advantages of refactoring are not in producing a better output, but rather in making the code run more efficiently, possibly using resources better, such as limited memory or processing power, and possibly making the code easier to read and thus easier to maintain in the future.  Refactoring is often used when the original code is asked to scale up its input to avoid run times that are unacceptable for the end users.

    While refactoring offers great upside, it comes with a few downsides and they are that it can be time consuming and add complexity to code, where a quicker solution provides results in acceptable time limits and the input isn't going to be scaled up in the future.  The efficiency of code tends to matter more as the input data grows.
    
### Application to Refactoring the Original VBA Script

    In this project we were asked to refactor the code to ensure that scaling up to thousands of stocks and multiple years wouldn't present an acceptable run time to achieve the results.  The original code was written in a simple logic that used  nested loops to read the entire dataset for each stock ticker, even though the data was already sorted by ticker and date. This meant that the spreadsheet was accessed more than necessary adding run time to the program.

    The refactored code was changed to read and separate all of the data with just one loop and load it into a series of arrays.  By accessing the outside spreadsheet just once, calculating the results, and moving it all to memory, we allowed the code to more quickly access results in memory and speed up the time for the program to run.  By reducing the number of reads to the outside data we took the effiency of the code from quadratic (n^2) using nested loops to linear (n) using a single loop and the results showed that reduction in run time.