# Stock Analysis Overview
The goal of this project was to use prior VBA code used to analysis Stocks for a given year, and refactor it the best we can to run faster. By doing this, if Steve would like to add more stocks to the data set, it will run much faster then before.

## Results Of Stocks
Firstly, if we looked at high performers in 2017 like SPWR, DQ, ENPH, and FSLR then on the surface these all would be great stocks to invest in. However given the 2018 data, we can see that the only stocks out of those 4 that were postive was ENPH. DQ was the worst performer in 2018 and given the high difference in year to year returns, it seems that DQ is a highly volitile stock to invest in. We can only recomend that ENPH, and RUN be invested in as they are the only two stocks to have positive returns year over year. If we could look at more statistics, we could see how volitile all the stocks are and pick a fairly even stock that has constant returns year over year for the parents.

<img src= "https://github.com/DAsInDavid1/stock-analysis/blob/main/2017_Returns.png" width=25% height=25%> 
<img src= "https://github.com/DAsInDavid1/stock-analysis/blob/main/2018_Returns.png" width=25% height=25%>

## Results Of Refactoring
The refactored script ran much faster as seen with these images.

### 2017 old vs new
<img src= "https://github.com/DAsInDavid1/stock-analysis/blob/main/VBA_Challenge_2017_Old.png" width=25% height=25%>
<img src= "https://github.com/DAsInDavid1/stock-analysis/blob/main/VBA_Challenge_2017.png" width=25% height=25%>

### 2018 Old vs new
<img src= "https://github.com/DAsInDavid1/stock-analysis/blob/main/VBA_Challenge_2018_Old.png" width=25% height=25%>
<img src= "https://github.com/DAsInDavid1/stock-analysis/blob/main/VBA_Challenge_2018.png" width=25% height=25%>

As seen in the images we improved the code by roughly a 1/2 second for both years. This will help largly for when Steve decicdes to add more tickers to the list. We would just need to update the arrays if he decides to add more tickers. 

## Summary of Refactoring
The advantage to refactoring the code is that it runs much faster and will save time in mineal tasks that Steve was going to do. We saw a great leap of 1/2 a second faster run time when we refactored, however it took me 3 hours to complete the refactoring and get the code to run correctly. If this refactoring saved time, then Steve would need to run the code 10,800 times for it to make up the 3 hours lost in refactoring. Steve could however, have his parents and freinds use the code, in which it would take less tiems for the code to be ran per person for the code to be worth refactoring. The refactoring would be advantagious when he wants to add ticker like said above, or if he puts multiple year of data into one worksheet. If he wanted to putthe 2017 and 2018 years together in the same workbook to see how the tickers are doing over the 2 year spand then the code would be running much faster then if we kept the old code. When refactoring we need to keep in mind how much time we are spending, to how much time can possibly be saved by doing it and think about if it is truley worth it in the long run.

