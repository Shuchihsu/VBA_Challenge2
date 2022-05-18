# Stcok Analysis with VBA Macro

## Overview of Project

  Steve would like to analyze the stocks for his parents for year 2017 and 2018 to find out how their profolios are doing. 
  We are using VBA macro and specifically with loop function to analyze the returns of each stock. We are also looking at the daily total volume.
  We would like to refactor the original macro to a more efficient macro so we are able to run a bigger data set in the future.
   

### Results

  We use the timer in the code so we can see if the refactor code is running faster (or slower) than the original code. 
  Below are the two screenshots for year 2017 and 2018 respectively. We can tell the Timer is showing around 0.08 sec for each year. 
  Red highlight is showing the return which is below 0 and green highlight is showing the return highler than 0.

![2017 runtime](https://github.com/jkmom/VBA_Challenge2/blob/main/Resources/VBA_Challenge_2017.png)


![2018 runtime](https://github.com/jkmom/VBA_Challenge2/blob/main/Resources/VBA_Challenge_2018.png)

### Summary

## The advantages and disadvantage of the refactoring code
    
    The advantage of the refactoring code is to improve the efficiency. Some codes work but don't mean they are good codes.
     We can always to improve codes if it is necessary. 
    Especially for a large database, sometimes a bad macro could freeze if the codes are not efficient for run a big database.

    Ideally, if a refactor code could have better performance but it also could also turn out interesting. 
    Sometimes we might have a good concept how to restructure the codes 
    but VBA might not have enough good functions to support the concept. Also, we also need to see if it is necessary to modify the code. 
    If the database is small, 
    like the one in our module, the difference between a couple of seconds might not worth to put time and effort to refactor the code.

## The advantages and disadvantage of the refactoring scripts and original scripts
   
   The original script is good and it runs pretty quick to provide the daily total volumes and returns just using two loops. 
   Overall the structure is healthy not wasting too much time going through the rows and tickers. 
    However, we do have space to improve the time that we can save on going through the rows.
   
   The advantage of the refactor codes is to more focus on looking at the ticker specifically when going through the rows. 
   We don't have to look at the tickers we don't count when going throug the whole row.
   We can skip the tickers before and just count the ticker through the ticker loop (tickerIndex = tickerIndex +1 )









