# stock-analysis
## Overview of Project

  The purpose of this project is to refactor code that was used analyze a list of stocks my client was looking at to provide advice
  to his parents to he could make the best reccomendation. Based on the inital data set the original program ran with out problems 
  but the data set was too small to provide the best investment recommendation. I was given the task to make the program run more efficiently
  to use less system resources this could potentially facilatate the use of a larger data set for analysis in the future.
  
## Results

###  The tickerIndex is set equal to zero before looping over the rows. 
        Dim tickerIndex As Integer
        tickerIndex = 0
### Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
### The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
        End If
### The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
        For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
          
        
    Next i
![green_stocks_2017](https://user-images.githubusercontent.com/89167531/133366009-ccb1d4c7-463b-4662-9e05-12d467d11048.png)
![green_stocks_2018](https://user-images.githubusercontent.com/89167531/133366010-128699f9-4e81-4d5b-a90e-b7bcb3912a8b.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/89167531/133366011-417cc752-73d7-4ddd-b722-aeca0ae7fb24.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/89167531/133366012-e5952fe8-bcbe-4882-a750-cd51a5cd1745.png)

## Summary

  I summary refactoring the code in this instance provided a result that took approx. 25% of the time for excel to do the calculation.
  One would be able to infer that the new code would be less likely to crash and would be able to handle much larger data sets. One of 
  problems with refactoring the original script is if you are the original author of the script having to look at the solution you already 
  solved with a critical eye to find the flaws in the program that can provide a challenge. However if you have a script that is too resouce 
  heavy or is taking too long to run providing the a more efficient way to solve the problem is necessary. If you have someone to look at the 
  code and provide input that would be advisable. The original code used loops to get the value for each stock to input into the spreadsheet 
  creating a variable that represents the array of stocks was more efficient.
