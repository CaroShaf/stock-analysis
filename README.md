# VBA GREEN STOCK ANALYSIS & REFACTORING

## Overview of Project: 
Steve loved the Excel workbook that we provided him for a dozen green stocks to analyze their performance.  We were able to provide him with a 
VBA macro to present key elements of his desired stocks in a user-friendly, easy to read format.  To expand on this for future analysis and to add other stock 
tickers, we optimized the code to easily accommodate additional stock tickers and years of data.  Furthermore, the code performs more efficiently and thus can
accommodate much more data and still run efficiently.  The data used in our analysis was provided by Freddy Hayes Consultants.  We can obtain more tickers and more 
years of data from this same source to expand Steve's analysis in the future.

## RESULTS

### 2017 & 2018 Stock Performance

[2017 Stock Performance & Execution Time] (https://github.com/CaroShaf/stock-analysis/blob/master/Resources/VBA_Challenge_2017.PNG)

[2018 Stock Performance & Execution Time] (https://github.com/CaroShaf/stock-analysis/blob/master/Resources/VBA_Challenge_2018.PNG)

### Refactored Script
 
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the 
refactored script.

#### Execution times
refactored 2017 .74
original 2017 .79

refactored 2018 .76
original 2018 .67


## SUMMARY

#There are advantages and disadvantages to refactoring the original script for this stock analysis application.  The advantage is that the code is cleaner
#and can be edited more easily to accommodate changes in the data files for increasing numbers of stock tickers (more rows) and years (more active sheets).  
#The disadvantage of this particular refactoring is that it is not fully refactored because the contents of the ticker array is hard coded, and not a dynamic
#array at this time.  Another disadvantage is that the execution times are not always less for the refactored script than for the original script.  It may
#depend on the hardware used, the number of times the script is run (and stored in memory) as well as after modifying to a dynamic array to accommodate a different
#number of tickers and years in the future.
