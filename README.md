# Green Stock Analysis

## Overview
An analysis of green energy stocks for the years 2017 and 2018.

### Purpose
An old friend Steve has recently graduated with his finance degree and his parents are his first clients. They have an interest in green engery so they want to invest their money in green energy stocks. They were specifically looking to invest in New Energy Corporation "DAQO" (ticker **DQ**)  so Steve has given me and Excel file with stock data to analyaze for twelve different green stock companies. I have written a VBA code that will automate the analysis and Steve will then be able to use it with any stocks with accuracy and refactored the code so it will run more efficiently.

#### Results
##### 2017
The initial analysis of the stock data proved that almost all stocks had a positive return rate in the year 2017. The only stock that had negative returns (lost money) was **TERP** with a return of -7.2%. **DQ** had a return rate of 199.4% in 2017, so it looks like a good investment for Steve's parents at this point in having the highest return rate of the twelve stocks in 2017. The stocks that had the next two highest return rates after DAQO were **SEDG** (184.5%) and **ENPH** (129.5%). The stocks that round out the bottom three of return rates were **RUN** (5.5%) and **AY** (8.9%), while these two stocks were among the lowest performing in 2017, they did still have a positive return rate, resulting in investors still making money.

![VBA_Challenge_2017.png](C:\Users\kathe\Desktop\Stock\resources\VBA_Challenge_2017.png)

##### 2018
The return rate for the green energy stocks however, were not as positive in 2018. Only two companies had a positive return rate. The only stocks whose returns were positive were **RUN** (84.0%) and **ENPH** (81.9%). **RUN** had a Total Daily Volume of 267,681,300 in 2017 and 502,757,100 in 2018. **ENPH**'s Total Daily Volume in 2017 was 221,772,100 and 607,473,500 in 2018. **DQ** was the worst performing stock in 2018 with a return rate of -62.6% and Total Daily Volume of 107,873,900 (compared to 2017's 35,796,200). The other two stocks in the bottom three were **JKS** (-60.5%) and **SPWR** (-44.6%).

![VBA_Challenge_2018.png](C:\Users\kathe\Desktop\Stock\resources\VBA_Challenge_2018.png)

#### **Example of original code:**
```
For i = 0 To 11

    ticker = tickers(i)
    totalVolume = 0

    '5) loop through rows in the data
      Worksheets(yearValue).Activate

            For j = 2 To RowCount

            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value

            End If

            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value

            End If

            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value

            End If

            Next j

    '6) Output data for current ticker
      Worksheets("All Stock Analysis").Activate

      Cells(4 + i, 1).Value = ticker
      Cells(4 + i, 2).Value = totalVolume
      Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1

Next i
```
The code ran the 2017 and 2018 stock figures in 1.023438 and 1.007813 seconds, respectively. It should be noted that this performance was achieved while only considering a limited dataset of 12 variables. As the size of the dataset increases, it is likely that the time required to run the code and nested loops will also increase, potentially straining system resources.

#### **Example of _refactored_ code:**
```

    Dim tickerIndex As Single
    tickerIndex = 0


    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrice(12) As Single
    Dim tickerEndingPrice(12) As Single

        For i = 0 To 11
        tickerVolumes(i) = 0

        Next i

        For i = 2 To RowCount

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrice(tickerIndex) = Cells(i, 6).Value

            End If

            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrice(tickerIndex) = Cells(i, 6).Value

        'If the next row’s ticker doesn’t match, increase the tickerIndex.
            tickerIndex = tickerIndex + 1

            End If

        Next i
```
This refactored code ran the 2017 stock figures in 0.1328125 seconds and the 2018 stock figures in 0.1484375 seconds. To improve its performance, I created additional arrays and refactored the code to use "For loops" instead of "Nested loops". As a result, the refactored code was able to process the dataset approximately 7 times faster than the original version. This optimization not only improves the speed of the code, but also makes it more scalable and efficient for larger datasets.


#### Summary

- What are the advantages or disadvantages of refactoring the code?
    Refactoring the code can offer several advantages, including the reduction of duplicated code, improved efficiency and performance, and enhanced maintainability. This allows for easier changes and updates to the code in the future. However, refactoring can also be a time-consuming process that requires careful attention to detail. I have personally experienced challenges with debugging syntax errors that arose from typos during the refactoring process.

- How do these pros and cons apply to refactoring the original VBA code?
    The benefits of reducing duplicated code by removing nested loops included improved performance and easier code maintenance. Specifically, the code ran approximately seven times faster. However, the process of removing the nested loops was time-consuming and required careful attention to avoid syntax errors.
