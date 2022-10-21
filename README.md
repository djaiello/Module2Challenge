# Refactoring VBA Stock Analysis Code

## OVERVIEW OF PROJECT

    Our original project was to design a stock analysis tool in Excel, using VBA, that would take a spreadsheet of daily stock figures for twelve stocks and calculate the annual volume and annual return for each stock.  In addition, the results would be formatted to separate stocks based on the annual return, providing an easy visualization for the client to make financial recommendations on.

    The purpose of this project is to scale up the capability of the original project to produce the same results for a greater number of stocks that the client may want to analyze.  In scaling up the potential number of inputs to use in this analysis, we are tasked with refactoring the original code to ensure that is runs more efficiently than the original.  

## RESULTS

    The stock analysis was performed using the original code and the refactored code on both 2017 and 2018 stock data for a total of 4 runs as shown below: 
### ORIGINAL Code for 2017 Stock Data
![image](https://user-images.githubusercontent.com/114360511/197105189-0efe45d8-8e3c-45a6-ad4b-5c7219b7289b.png)

### REFACTORED Code for 2017 Stock Data
![image](https://user-images.githubusercontent.com/114360511/197105322-3b82b8e4-8c29-42f6-b990-e32411d7ce25.png)
       
### Result Summary for 2017
    The original code ran in 0.902 seconds while the refactored code ran in 0.156 seconds.  Thus the refactored code was much faster.
    
    
### ORIGINAL Code for 2018 Stock Data  
![image](https://user-images.githubusercontent.com/114360511/197105433-4836cf45-d132-44a1-af3a-e78dc65cbd49.png)

### REFACTORED Code for 2018 Stock Data
![image](https://user-images.githubusercontent.com/114360511/197105536-e841bd43-5633-4068-a08f-9caf1cf6e508.png)

### Result Summary for 2018
    The original code ran in 0.699 seconds while the refactored code ran in 0.132 seconds.  Thus the refactored code was much faster.


### Original and Refactored Code Link:



## SUMMARY

### Advantages and Disadvantages of Refactoring Code
    Refactoring is an important step in any code design of an algorithm.  The advantages of refactoring are not in producing a better output, but rather in making the code run more efficiently, possibly using resources better, such as limited memory or processing power, and possibly making the code easier to read and thus easier to maintain in the future.  Refactoring is often used when the original code is asked to scale up its input to avoid run times that are unacceptable for the end users.

    While refactoring offers great upside, it comes with a few downsides and they are that it can be time consuming and add complexity to code, where a quicker solution provides results in acceptable time limits and the input isn't going to be scaled up in the future.  The efficiency of code tends to matter more as the input data grows.
    
### Application to Refactoring the Original VBA Script

    In this project we were asked to refactor the code to ensure that scaling up to thousands of stocks and multiple years wouldn't present an acceptable run time to achieve the results.  The original code was written in a simple logic that used  nested loops to read the entire dataset for each stock ticker, even though the data was already sorted by ticker and date. This meant that the spreadsheet was accessed more than necessary adding run time to the program.

    The refactored code was changed to read and separate all of the data with just one loop and load it into a series of arrays.  By accessing the outside spreadsheet just once, calculating the results, and moving it all to memory, we allowed the code to more quickly access results in memory and speed up the time for the program to run.  By reducing the number of reads to the outside data we took the effiency of the code from quadratic (n^2) using nested loops to linear (n) using a single loop and the results showed that reduction in run time.
