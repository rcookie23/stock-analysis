# stock-analysis

## Overview of Project

In this project, we were asked to analyze data on 12 different stocks over the course of two years. Specifically, we were asked to determine the total daily volume and the percentage of return for each of the 12 stocks over the course of the years 2017 and 2018.  In order to performed this analysis, we created VBA script in Excel to comb through the large data sets containing stock information. 

## Results

### Stock Performance

The images below display the data we obtained on stock performance for the years 2017 and 2018. In 2017, every stock had a positive return except for TERP. The year 2018 told a much different story, as all stocks had a negative return except for ENPH and RUN. The charts also show that there was much variation in the total daily volume of each of the stocks. Some stocks experienced an increase in volume from 2017 to 2018, and some experienced a decrease. 

![Stock Results 2017](/Resources/results_2017.png)

![Stock Results 2018](/Resources/results_2018.png)

### Building the VBA Script

For this analysis, we wrote two different VBA scripts to deliver the relevant stock data. The first script was sufficient to give us the information we wanted, but the code was not as efficient as we would have liked it to be. So, we refactored the script to make it more efficient. The differences between the two scripts is explained below. 

#### Original VBA Script

Both the original and the refactored script used the array below to account for the ticker symbols of the 12 different stocks:

![Tickers Array](/Resources/tickers_array.png)

In the original VBA script we created, we used a nested For loop to loop through above array. For each element of the array, the loop combed through the entire data set looking for the parameters set out in the 3 if-then statements in the inner loop. Once the entire loop was complete for one ticker, the process repeated, looping through the entire data set again for the next ticker. The relevant portion of the original script is displayed below:

![Original Script](/Resources/original_loops_outputs.png)


For this script to run, it had to loop through the entire data set 12 times, one time for each ticker. This method of combing through the data set was not as efficient as we would have liked, so we refactored the script to increase its efficiency.

#### Refactored VBA Script

In the Refactored code, we set up 3 new arrays (in addition to the index array) instead of using a nested For loop. Instead of making startingPrice, endingPrice and totalVolume standard variables, we set them up as arrays with 12 elements each. We set up a variable called tickerIndex, whose purpose was to account for which ticker was currently being analyzed. The refactored script only had to comb through each data set one time. The script was written to add up the total daily volume for each ticker until it realized that it had reached the end of that ticker's data. Once it reached the end of a ticker's data, the script recorded the final closing price for that ticker and then told the tickerIndex variable to increase by 1. This caused the summation of total daily volume for the next ticker in the array to began. The script then recognized that a new ticker was being analyzed, and it recorded the initial closing price for the new ticker. The relevant portions of the refactored script is displayed below:

![Refactored Loops](/Resources/refactored_loops.png)
![Refactored Outputs](/Resources/refactored_outputs.png)


#### Time to Complete Script

As described above, the refactored script was much more efficient that the original script, and thus had a much quicker run time. Compare the run times of the two scripts below: 

##### Original Script Run Times

![2017 OG Time](/Resources/2017_OGtime.png)
![2018 OG Time](/Resources/2018_OGtime.png)

##### Refactored Script Run Times

![2017 Refactored Time](/Resources/2017_Refactored_time.png)
![2018 Refactored Time](/Resources/2018_Refactored_time.png)

## Summary

As evidenced above, taking time to refactor code can lead to great efficiencies in the code. It can decrease the time it takes for the code to run, and it can eliminate unnecessary redundancies in the code. This can be a great advantage, especially if you are running a complicated code that required lots of time and resources to complete. A disadvantage to refactoring code can be summed up by the old idiom, "if it ain't broke, don't fix it." If a code works and gives you the correct results, it may be a waste of time and resources to refactor the code to make it more efficient. You also run the risk of mangling your original code beyond recognition (always save a copy of the original code!). 

In our case, refactoring the code allowed the code to run faster, However, it only took the original code a fraction of a second to run. The refactored code, although significantly faster in terms of percentage, was only a fraction of a second faster than the original. However, if the size of the dataset was drastically increased (e.g., if we were analyzing every stock in the market instead of just 12 stocks), then the benefits of making the code more efficient would be more readily apparent. 

