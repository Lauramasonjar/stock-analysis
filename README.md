# Module 2 Challenge: VBA Challenge

## Overview

In this challenge, I edited, or **refactored**, the Module 2 solution code to loop through all the data for 2017 and 2018 a single time to collect the same information that I did throughout this module. Then, I ascertained whether refactoring my code successfully made the VBA script run faster than it's previous run time. 

### Purpose

The purpose of this challenge is to help Steve expand the dataset from the previous workbook to include the entire stock market over the last few years and to run it more efficiently.

## Analysis and Challenges:

To analyze the stock information for 2017 and 2018 in the fastest, most efficient way possible, I utilized Excel and VBA code.  In VBA I wrote a subroutine for "All Stocks Analysis" and a subroutine for "All Stocks Analysis Refactored".  To get my results I used the following data:
	1. Original script from "All Stocks Analysis"
	2. Challenge Starter Code
	3. "All Stocks Analysis Refactored"

#### Data Background
The subroutines for "All Stocks Analysis" and for "All Stocks Analysis Refactored" used complex code that included arrays, loops, a timer and data formatting.  For "All Stocks Analysis" data was being pulled from 12 different stocks on ticker value, date issued, opening, closing, adjusted closing price and their volumes for 2017 and 2018.  For All stocks Analysis Refactored" the same data was being pulled, but this time to include the entire stock market for those two years.  The timer is the code that provided us with a visual result for the efficiency of our refactored subroutine in comparison to the subroutine for "All Stocks Analysis".  Refer to images below for the refactored subroutine.

![2017 Spreadsheet](https://raw.githubusercontent.com/Lauramasonjar/stock-analysis/main/2017%20Spreadsheet.png)

![2017 Run Time MsgBox](https://raw.githubusercontent.com/Lauramasonjar/stock-analysis/main/VBA_Challenge_2017.png)

![2018 Spreadsheet](https://raw.githubusercontent.com/Lauramasonjar/stock-analysis/main/2018%20Spreadsheet.png)

![2018 Run Time MsgBox](https://raw.githubusercontent.com/Lauramasonjar/stock-analysis/main/VBA_Challenge_2018.png)




#### Analysis

Below is the starter code provided for the refactored subroutine:




  Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    

    '1b) Create three output arrays   
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero. 
    
        
    ''2b) Loop over all the rows in the spreadsheet. 
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            

            '3d Increase the tickerIndex. 
            
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
## Summary:
You should write code that's readable and efficient while producing accurate outcomes.  The advantage of refactored code is that is provides a faster outcome whereas the disadvantage is that it takes longer to write. Overall, you can avoid wasting time later if you spend the time in the beginning writing readable, clean, well- organized code.
