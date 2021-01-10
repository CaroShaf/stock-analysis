# VBA GREEN STOCK ANALYSIS & REFACTORING

## Overview of Project
Steve loved the Excel workbook that we provided him for a dozen green stocks to analyze their performance.  We were able to provide him with a 
VBA macro to present key elements of his desired stocks in a user-friendly, easy to read format.  To expand on this for future analysis and to add other stock 
tickers, we optimized the code to easily accommodate additional stock tickers and years of data.  Furthermore, the code performs more efficiently and thus can
accommodate much more data and still run efficiently.  The data used in our analysis was provided by Freddy Hayes Consultants.  We can obtain more tickers and more 
years of data from this same source to expand Steve's analysis in the future.

## RESULTS

### 2017 & 2018 Stock Performance

[2017 Stock Performance & Execution Time] (https://github.com/CaroShaf/stock-analysis/blob/master/Resources/VBA_Challenge_2017.PNG)

[2018 Stock Performance & Execution Time] (https://github.com/CaroShaf/stock-analysis/blob/master/Resources/VBA_Challenge_2018.PNG)

Looking at the stock performance of 2017, it is easy to see why Steve's parents were interested in DQ stocks, which had an almost 200% return.  
In 2018, however, they were down about 63%.  It may be worthwhile to see if there was an initial excitement for new stock options or a new product in 2017 with
some sort of economic challenge in 2018 that caused the sharp decline.  Overall, all but two of the dozen green stocks examined performed poorly between 2017 and
2018.  The companies that continued to have positive returns were ENPH and RUN.  Steve may want to look at more years of data for all of the stocks before
recommending DQ or even ENPH and RUN.

### Refactored Script
 
The images referenced above show the refactored execution times for 2017 and 2018.  The original script's execution times are not depicted in a pic, but recorded
below the refactored times as a comparison.

#### Execution times
Refactored 2017 .76

Original 2017 .79

Refactored 2018 .75

Original 2018 .67

The primary modification of the basic script was to include storage of data in arrays before performing calculations and outputing results, so as to loop through the
data set in a more systematic method, requiring fewer loops.  As predicted the change in script affects the execution times, but not always in predictable ways.  It
can be observed that the refactored script for 2017 was slightly faster while the refactored script for 2018 was slower.

## SUMMARY

There are advantages and disadvantages to refactoring the original script for this stock analysis application.  The advantage is that the code is cleaner
and can be edited more easily to accommodate changes in the data files for increasing numbers of stock tickers (more rows) and years (more active sheets).  
The disadvantage of this particular refactoring is that it is not fully refactored because the contents of the ticker array is hard coded, and not a dynamic
array at this time.  Another disadvantage is that the execution times are not always less for the refactored script than for the original script.  It may
depend on the hardware used, the number of times the script is run (and stored in memory) as well as after modifying to a dynamic array to accommodate a different
number of tickers and years in the future.
