# VBA of Wall Street
## Overview of Project

### Purpose

In a general sense, the purpose of this project is to familiarize ourselves with the language of coding, particularly in VBA, by using an application that is familiar to us, Excel. In this challenge, we are helping analyze a dataset of stocks to find each unique stock's total volume for the year and the percentage of return that stock gave in the last year so that Steve can help his parents make smarter investment choices. In the challenge, we took the code that we had worked on in the module, which allowed us to do everything we needed, and refactored, or restructured, the code to make it more efficient in how quickly it would run the code and output the results. The purpose of doing so is so that Steve can utilize the refactored code to expand the dataset he is looking at over the entire stock market. This would mean thousands of more stocks would need to be analyzed and this refactored code would likely output results much quicker than if we were to use the non-refactored code.

## Results

In our table that is created in Excel using our VBA code, we see that for 2017 and 2018 there are only two stocks that had positive returns: ENPH and RUN. We can conclude that these two stock would be strong choices to invest in as they seem the least risky. TERP is the only stock to be in the red both at the end of both years, so it would be best to stay away from it.

In terms of our code, we can see in the images below that our refactored code ran in under a fourth of a second. To build the same table and receive the same results using our unrefactored code, I saw that it took a little more than 3 seconds each time the code was run making our refactored code significantly faster and more efficient. The most important part of the refactored code that made it more efficient is seen here:

```
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
```

![VBA_Challenge_2017.png](https://github.com/amirshirazi1/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![VBA_Challenge_2018.png](https://github.com/amirshirazi1/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

1. What are the advantages or disadvantages of refactoring code?

2. How do these pros and cons apply to refactoring the original VBA script?
