# Challenge 1: VBA of Wall Street

## Overview of Project
Analyse the stock data for 12 stock tickers for the years 2017 and 2018 using a Visual Basic For Applications macro. 

### Purpose
To analyze the stock dataset to generate the annual volume and return rate for a chosen calendar year. The annual volume is the sum of the number of shares traded each day, and the return rate is the percentage difference in price from the beginning of the year to the end of the year. The implementation is also being assessed for performance and refactored to improve runtime.

## Results

### Results of Stock Analysis for 2017 and 2018
The VBA_Challenge.xlsm workbook includes the worksheet "All Stocks Analysis" and the VBA script to perform the analysis. [VBA_Challenge.xlsm] (path)

These images represent the results of running the VBA script for 2017 and 2018.

![2017 Results] (path)
![2018 Results] (path)          
        
### Performance Results of Stock Analysis Before Refactoring 
   The screenshots below display the run times for the intial implementation of the stocks analysis script.
   
   ![hello!](~./images/adam-solomon-hello.jpg~ "adam solomon's hello”)

### Performance Results of Stock Analysis After Refactoring 
  The code was refactored to iterate over the stock data once and record the values in arrays. The screenshots below display the run times for the refactored implementation of the stocks analysis script.
    
    ![Outcomes_vs_Goals](path)


As a result of the refactoring, performace runtime was improved by 82% for 2017 and by 80% for 2018.   

## Summary

- What are the advantages or disadvantages of refactoring code?
 
 One of the main advantages of refactoring code can be an improvement in run time performance. Another advantage can be improved readability and maintainability. A key disadvantage is that the solution must be re-tested as errors could be introduced during the process of refactoring. In some cases, refactoring can actually make the code more complex and harder to maintain if the developer tries too hard to minimize the amount of code.    

- How do these pros and cons apply to refactoring the original VBA script?

 By iterating through the stock data once, rather than 12 times, the performance was significantly improved. The algorithm was made somewhat more complex in determining when a new stock ticker started and ended. The additional complexity was managed through descriptive comments to ensure future maintainability was not compromised. 

## Additional Notes

The refactoring focused primarily on runtime performace improvements. The current implementation makes several assumptions about the data: 

 * The data is ordered by stock and date 
 * There will always be data for every stock ticker
 * The user running the script will enter a year for which data is available (no error checking).

 Additional refactoring and code improvements could be implemented to make the solution more resilient and flexible. 