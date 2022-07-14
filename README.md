# stock-analysis
Analyzing DAQO and other green energy stocks

## 1.	Overview of Project
1.	The clients of this project want to invest in green energy, specifically, they want to invest all of their money into DAQO New Energy Corporation (DQ stock). The purpose of this project is to evaluate DAQO stock along with 11 other different green energy stocks to determine which are the best stocks to invest in. Specifically, the analysis of these stocks looked at two factors. The first considered the total daily volume of each stock. While, the second factor analyzed the return percentage rate, verifying whether it was a positive or negative return. 

## 2. Results

### Stock Performance

Graphs sorted alphabetically by stock name

![2017 stock return](/Resources/2017_stock_return.PNG)
![2018 Stock Return](/Resources/2018_stock_return.PNG)

While comparing the stock performance of the 12 green energy stocks, seen in the images above (sorted alphabetically by stock name), the first thing that it easily noticed is which stocks had a positive return and which stocks had a negative return. In 2018, only two stocks yielded a positive return, ENPH and RUN. Coincidentally, those two stocks also yielded a positive return in 2017. 

Graphs sorted by total daily value (descending order)

![TDV_sorted_2017.PNG](/Resources/2017_stock_return.PNG)
![TDV_sorted_2018.PNG](/Resources/2018_stock_return.PNG)

Secondarily, the next comparison this is considered is the total daily volume. In 2018, we see that the three stocks that yielded the highest total daily volume are ENPH, SPWR, and RUN respectively. 

In 2017, we see that SPWR did yield the highest total daily volume, while ENPH and RUN are still in the top 5. 

Comparing all the data discussed above, I would suggest that the stock with the most promising results would be ENPH as it has a positive return for the past two years, while also earned the top volume in 2018 and in the top 5 for 2017. 

My second suggestion would be RUN stock as it also had positive returns for the past year and in the top four of earning the most total daily volume. 

The clients of this project specifically wanted to invest in DQ stock. While we can see that it had close to a 200% return in 2017, it had a large negative return in 2017, at -62.6%. We can also notice that in 2018 it earned the lowest total daily volume and in the bottom three in 2017. From this data, I would urge the clients to consider to diversify their money, at the very least instead of investing all of their money in DAQO. I would encourage them to consider investing, at a minimum, part of their money in ENPH or RUN stock.  


### Coding Execution Times

In the original VBA code, as seen in the graph below, the code ran through the each ticker (each stock) and then calculated the criteria for each stock (total volume, starting price for each year, and ending price for each year). This means it ran through all of the rows of the worksheet 12 different times to first consider what stock it was evaluating, and then calculating those three values for each stock/ticker. 
![OriginalCode.PNG](/Resources/OriginalCode.PNG)

In the reformatted VBA code, seen below, it only had to run through all the rows of the worksheet once. While running through the rows, it sorted the tickers/stocks into arrays, and then evaluated the same three necessary data values as the original code. Since the code only had to run through all of the rows of the worksheet once, then it decreased the time to execute the calculations substantially.
![Reformated_code.PNG](/Resources/Reformated_code.PNG)

## Summary

As seen in the images below, an advantage of reformatting code is the substation decrease of time it took for the program to run through the necessary calculators. On the left images, we can see that the original VBA code took somewhere between .6 and .7 seconds to make all the necessary calculations. On the right, we can see that the reformatted code took less than .1 second to make the same calculations. 

However, one disadvantage of reformatting code, is the time it takes for the programmer to update the computer program. That is, if is even possible to rewrite and make it better. 

![2017_original_time.PNG](/Resources/2017_original_time.PNG)
![2017_refactored_time.PNG](/Resources/2017_refactored_time.PNG)

![2018_original_time.PNG](/Resources/2018_original_time.PNG)
![2018_refactored_time.PNG](/Resources/2018_refactored_time.PNG)

One pro of refactoring this VBA code is that it decreased the time for the code to run by approximately 85%. However, noticing that the original code made the necessary calculations in under a second, it may not have been worth the time to refactor the code to make such a minimal decrease. 




