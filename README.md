# VBA_Challenge

## Overview of Project:This project is an analysis of returns for the entire stock market for years 2017 and 2018. 
Results: Stock performance from 2017 to 2018. Overall stocks performed better in 2017. 
Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    'Looping through the Return values and change to green if positive value otherwise red
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If

Stock return results from 2017 analysis

![pic1](https://github.com/Klubbers0/VBA_Challenge/blob/main/Resources/StockPerformance2017.PNG)

Stock return results from 2018 analysis

![pic2](https://github.com/Klubbers0/VBA_Challenge/blob/main/Resources/StockPerformance2018.PNG)


![pic5](https://github.com/Klubbers0/VBA_Challenge/blob/main/Resources/VBA_challenge_2017.PNG)

Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?

