# VBA_Challenge
**Module 2 Challenge Written Analysis**

**Project Overview**

The purpose of this project was to refactor our stock analysis code. Previously to this challenge we had written code in VBA to calculate the total volume and returns for a set of stocks. In this challenge we rewrote a good portion of that VBA code to make it more efficient and run quicker. We made 3 additional arrays for tickerVolumes, tickerStartingPrices, and tickerEndingPrices and we used an additional variable, tickerIndex, to make sure we are referencing the correct ticker during our loop.

**Results**

The returns in 2017 were substantially better than the returns in 2018. The only 2 companies that outperformed in 2017 were RUN and TERP. The sum of all total daily volumes was fairly similar from 2017 to 2018 with slightly more volume in 2018. The execution times in the refactored script was significantly faster. Almost 8 times faster. 

![Run Time 2018](https://user-images.githubusercontent.com/95661553/148697837-f6f2faca-66b4-4d7f-bb8c-00ed07ff3685.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/95661553/148697844-9da9915d-ae3a-48c2-b89b-61aff4c36b48.png)


The reason for this is because instead of running through all the data for each ticker, it just runs through the data once. The code goes to the next ticker once it has gone through all of the instances of the first ticker. We used the indexTicker variable to do this. Once there is a new ticker we add 1 to the tickerindex which effectively looks at the next ticker in the list.

If Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
                            tickerindex = tickerindex + 1
End If   

A potential negative to this new refactored code is that if the tickers on the data pages were not in the same order as the ticker in our code, this code would not work. Our old code would work regardless of order of tickers in our code.

**Summary**

**Advantage and Disadvantage Statement:**

The main advantage to refactoring code is that you can make your code more efficient and work better; however, the disadvantage is that it can take substantial time to refactor code that already works.

**Statement applies to our script:**

In this circumstance we sped up the run time significantly by only going through the data once, but even though we made the code 8 times faster, we spent siginificant time getting the overall run time from .8 to .1 seconds which in nominal terms is not that big of a difference. 

