# VBA_Challenge

## Overview of Project:This project is an analysis of returns for the entire stock market for years 2017 and 2018. 
Results: Stock performance from 2017 to 2018. Overall stocks performed better in 2017. 




    dataRowStart = 4
    dataRowEnd = 15

    'Looping through the Return values and change to green if positive value otherwise red
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If

### Stock return results from 2017 analysis

![pic1](https://github.com/Klubbers0/VBA_Challenge/blob/main/Resources/StockPerformance2017.PNG)

Stock return results from 2018 analysis

![pic2](https://github.com/Klubbers0/VBA_Challenge/blob/main/Resources/StockPerformance2018.PNG)

### Comparing the execution times from the original code to the refractored code. 


Original execution time for 2017

![pic3](https://github.com/Klubbers0/VBA_Challenge/blob/main/Resources/ss%202017%20Green%20Stocks.PNG)

Refactored execution time for 2017

![pic4](https://github.com/Klubbers0/VBA_Challenge/blob/main/Resources/VBA_challenge_2017.PNG)

Original execution time for 2018

![pic4](https://github.com/Klubbers0/VBA_Challenge/blob/main/Resources/SS%202018%20Green%20Stocks.PNG)

Refactored execution time for 2018

![pic5](https://github.com/Klubbers0/VBA_Challenge/blob/main/Resources/VBA_Challenge_2018.PNG)


## What are the advantages or disadvantages of refactoring code?
### The advantage

-Code can be cleaned up to be less chaotic and more logical and comprehensive. 

-It is easier to spot logical errors in the code that contains nested and loops.
### The disadvantage

-Refactoring the code can affect the end test outcomes

-When code is complex, it is best to spilt into several sub macros. 

## How do these pros and cons apply to refactoring the original VBA script?
When code is clean, ordered, free of errors, and contains comment marks it is easier to modify and maintain in the future.



