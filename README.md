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

Using the `for loop` and the `if-then` statements make iterating through each of the `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices` much more efficient. But particularly, I think the part that made the code more efficient was the use of `tickerIndex` and the final `if-then` statement that increased the `tickerIndex` if the next ticker was a new stock. In our previous code, this was not included. We had used a nested `for loop` which iterated over every ticker and row in the sheet each time a new ticker was being accessed and assessed. Our refactored code only iterated over each ticker once and then moved on to the next one without including every single row in the sheet as it kept going.

![VBA_Challenge_2017.png](https://github.com/amirshirazi1/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![VBA_Challenge_2018.png](https://github.com/amirshirazi1/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

1. What are the advantages or disadvantages of refactoring code?
    
    I'll begin with the advantages: the code can run much more efficiently and is much easier to read. Refactored code is also great for initializing over large quantities of data because of how much more efficiently it can run. A disadvantage of refactoring code is that if there is one minor mistake in the code then it is a little more challenging to find. Additionally, to refactor code, you have to completely understand precisely what every line of code is doing because if you do not, then there is an increased chance of error and a much higher chance that you will break the code.
    
2. How do these pros and cons apply to refactoring the original VBA script?

    As I mentioned above, one of the pros of refactoring the script is that the time it took for the original VBA script to run was upwards of 3 seconds. When the refactored code was used, it only took .10 seconds for both years to run the VBA script.
    
    I ran into my issues when refactoring the code that contributed to what I viewed to be one of the biggest cons: minor mistakes that break the code. For example, in my `if-then` statement for checking the current row against a different row, I forgot to label a column in my line so it was written as `If Cells(i, 1).Value <> Cells(i - 1).Value Then` which created an error for the entire `for loop` that this line was running in. I agonized over trying to figure out what the issue was before or immediately after the `for loop` that was being highlighted as the error. But finally, I read over the rest of my code and noticed that I hadn't defined the column. Once I fixed that, the rest of the code ran perfectly. Because as the original VBA script runs each section of code in a way that is slightly less dependent on the previous lines the room for this type of large-scale code-breaking is less likely to happen than in the refactored.
    
    However, in the original VBA script, it is much easier to write and keep track of everything when your data set is smaller (as in the dozen tickers in this challenge), but if we were to use the same code for the thousands of stock in the stock market it would be arduously difficult and lengthy to write. This is why refactoring code is useful and why it should be done.
