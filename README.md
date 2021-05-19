# VBA of Wall Street

## Overview of Project
In this challenge we are given a scenario where we have a friend named Steve whose parents are interested in investing in Green Energy Stocks. We are tasked with analyzing a stock of their choice DAQO and other Green energy stocks to see if the stocks are a good investment or not. Using some basic VBA macros and stock analysis for yearly return we have created and optimized a Macro for analyzing stocks given specific parameters of a stock. 

### Purpose
The main purpose of this challenge is to develop a deeper understanding of VBA and the utilization of Macros to perform simple analysis of stock prices. With a better understanding of statistics and stock market indicators (doji,EPS,P/E,etc.) VBA can become a more powerful tool with even greater analysis than what we did. In the challenge specifically we had to optimize our code to have a faster runtime. 

## Analysis and Challenges
The challenge of this week was to refactor our code that we created from the module. We had to optimize the runtime which is important for real life scenarios where we will have many more data points than in our module. 

### Analysis of Original vs. Refactored
In the module we utilized a nested for-loop to loop through an array of cells in the worksheet to be analyzed. We had a defined array of tickers we were looking for, then we looped through each ticker and each row until the end of the data set to add the daily volumes and find the starting and last closing price to calculate the yearly return. From college I learned that the BigO notation for for-loops is O(n) where n is the number of iterations needed to be calculated. Assuming 1 sequence is 0(1), we have to iterate N times and if we have another for-loop, say M, the time complexity is exponential N * M times N^2. This gives you O(N^2). See below for the original VBA code:

	Sub AllStocksAnalysis()
	    'Calling in yearValue subroutine
	    Dim yearValue As Variant
	    Dim startTime As Single
	    Dim endTime As Single
	        
	    yearValue = yearValueAnalysis()
	    startTime = Timer
	    Worksheets("All Stocks Analysis").Activate
	'Creating Headers
	    Cells(1, 1).Value = "All Stocks (" + yearValue + ")"
	    Cells(3, 1).Value = "Ticker"
	    Cells(3, 2).Value = "Total Daily Volume"
	    Cells(3, 3).Value = "Return"
	'Create an arry to store all the tickers
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
	'change to sheet we will be doing our analysis in
	    Worksheets(yearValue).Activate
	    totalVolume = 0
	    rowStart = 2
	    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
	    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
	    Dim first_price As Double
	    Dim last_price As Double
	    
	    For i = 0 To 11
	    'Everytime we loop in the outer loop we want to be back in the sheet we want to do analysis in
	        Worksheets(yearValue).Activate
	        totalVolume = 0
	        'looping through the tickers
	        ticker = tickers(i)
	        
	        For j = rowStart To rowEnd
	        'increase totalVolume if ticker is "DQ"
	            If Cells(j, 1).Value = ticker Then
	                totalVolume = totalVolume + Cells(j, 8).Value
	            End If
	        'Checks if the current row is the First Row of DQ's Data, if so then set current closing price as the first price
	        ' Does this by having condition of cell = DQ and there is no previous DQ then it has to be first closing price
	            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
	               first_price = Cells(j, 6).Value
	            End If
	        ' checks if current row is DQ and if there are no other DQ after it
	            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
	               last_price = Cells(j, 6).Value
	            End If
	        Next j
	        'take the result from loop j and print it out in the All Stock Analysis sheet
	            Worksheets("All Stocks Analysis").Activate
	            Cells(i + 4, 1).Value = ticker
	            Cells(i + 4, 2).Value = totalVolume
	            Cells(i + 4, 3).Value = (last_price / first_price) - 1
	    Next i
	    
	    endTime = Timer
	    MsgBox "This code ran in " & (endTime - startTime) & "Seconds for the year " & (yearValue)
	        
	End Sub


	Sub formatAllStocksAnalysisTable()
	'Formatting
	Worksheets("All Stocks Analysis").Activate
	'Bolds
	Range("A3:C3").Font.Bold = True
	'bottom line
	Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
	Range("C4:C15").Borders(xlInsideHorizontal).LineStyle = xlContinuous
	'Italic
	Range("A3:C3").Font.Italic = True
	' used to disply only a significant digit with commas
	Range("B4:B15").NumberFormat = "#,##0"
	' used to convert to percentage
	Range("C4:C15").NumberFormat = "0.00%"
	Columns("B").AutoFit
	'Defining your range
	Dim returnRange As Range
	Set returnRange = Range("C4:C15")
	    For Each cell In returnRange
	    ' if negative then turn the font red
	        If cell < 0 Then
	            cell.Font.ColorIndex = 30
	            cell.Interior.ColorIndex = 15
	    'else if positive then green
	        ElseIf cell > 0 Then
	            cell.Font.ColorIndex = 10
	            cell.Interior.ColorIndex = 19
	    'else clears it
	        Else
	        cell.Interior.Color = xlNone
	        End If
	    Next

	End Sub

	Sub checkers()
	Worksheets("All Stocks Analysis").Activate
	Dim counter As Integer
	counter = 0

	    For i = 1 To 8
	        For j = 1 To 8
	            If counter Mod 2 = 0 Then
	            Cells(i + 3, j + 6).Interior.ColorIndex = 1
	            Else
	            Cells(i + 3, j + 6).Interior.ColorIndex = 2
	        End If
	        counter = counter + 1
	       Next j
	        counter = counter + 1
	    Next i
	            

	End Sub

	Sub ClearWorksheet()

	    Cells.Clear

	End Sub

	Function yearValueAnalysis()
	yearValue = InputBox("What year would you like to run the analysis on?", "Enter 2017 or 2018", 2018)
	yearValueAnalysis = yearValue
	End Function



In the challenge we must take the gained knowledge and refactor our code to lessen the time. As explained previously a for-loop has time complexity of O(n). For the challenge we had to create multiple for-loops and store the data into arrays. I would have never thought of this method and was extremely intrigued and excited to see such a cool work around for increasing performance. I've ran into this common algorithm problem before in the binomial coefficient problem. In this problem, if we have nth terms, our runtime will be n^n which is poorly written code. From binomial expansion we see that our next coefficient will always be (n-1) meaning we have already calculated the term previously. With our VBA code, we took the values from our for-loop and stored them into an array that was already created. In recursive programming you take the previous value and add it to the new value having to recalculate the first value then adding it to the new calculated value. This repetition is the opposite of what coding is supposed to be for. There is a term called dynamic programming where have the previous data is stored and can be used in the current calculation to reduce repetition. In our case we created empty arrays and stored our single for-loop values into that specific array. Since single for loops are O(n) having multiple for-loops would only be N * O(n) which would give us a shorter time than something exponential. See below for the array implementation. 

    Sub AllStocksAnalysisRefactored()
        Dim startTime As Single
        Dim endTime  As Single

        yearValue = InputBox("What year would you like to run the analysis on?", "Enter 2017 or 2018", 2018)

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
        tickerIndex = 0

        '1b) Create three output arrays
        Dim tickerVolumes() As Long
        Dim tickerStartingPrices() As Single
        Dim tickerEndingPrices() As Single
        
        ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For t = 0 To 11
           ReDim tickerVolumes(t)
           ReDim tickerStartingPrices(t)
           ReDim tickerEndingPrices(t)
           tickerVolumes(t) = 0
        Next t
        
        'Looping through the tickers
        ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
            '3a) Increase volume for current ticker index
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

            '3b) Check if the current row is the first row with the selected tickerIndex.
            'If  Then
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                End If
                
            'End If
            
            '3c) check if the current row is the last row with the selected ticker
             'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
            'If  Then
                
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                   tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                '3d Increase the tickerIndex by 1 to 11.
                    tickerIndex = tickerIndex + 1
                End If
        Next i
        
        '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For j = 0 To 11
            
            Worksheets("All Stocks Analysis").Activate
                Cells(j + 4, 1).Value = tickers(j)
                Cells(j + 4, 2).Value = tickerVolumes(j)
                Cells(j + 4, 3).Value = (tickerEndingPrices(j) / (tickerStartingPrices(j)) - 1)
            
    Next j
        
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


### Challenges and Difficulties Encountered
There were many challenges and difficulties that I created for myself. I initially wanted to make this Macro dynamic and scale with nth amount of ticker symbols. I tried implementing a tickers function that kept giving me an overflow error, so I completely removed that function. I also ran into multiple overflow errors when trying to keep the tickerVolumes, tickerStartingPrices, tickerEndingPrices dynamic. This ended up working by using a separate loop and using the ReDim function that allows you to resize the dimensions of your array. Another issue I ran into was trying to create a function to store a user's input. I was trying to create a button that asked for the year and when inputting the year, you can click the analysis button and get our analysis for that current year. For some reason I had trouble passing the function's stored data value into another Macro. In the original I created a function that called itself but that is just repetitive. All in all, to make this modular to work for any stocks and any amount of data was the limitation of my knowledge although I learned more than I needed researching these different functions and errors.  

## Results
From the results we can clearly see that the Refactored code is a lot faster than the nested for-loop code. See below for a comparison. The main takeaway from these results is that we want to avoid any code that will take an exponential amount of time to iterate through the loops, but at the same time make it as robust as possible. 

![Original Nested 2017 Runtime](https://github.com/lo7kyle/stock-analysis/blob/main/Resources/Nested%20Loop%202017_Time.PNG)
*Fig. 1: Original Nested 2017 Runtime
![Refactored 2017 Runtime](https://github.com/lo7kyle/stock-analysis/blob/main/Resources/Refactored%202017_Time.PNG)
*Fig. 2: Refactored Nested 2018 Runtime

![Original Nested 2018 Runtime](https://github.com/lo7kyle/stock-analysis/blob/main/Resources/Nested%20Loop%202018_Time.PNG)
*Fig. 3: Original Nested 2017 Runtime
![Refactored 2018 Runtime](https://github.com/lo7kyle/stock-analysis/blob/main/Resources/Refactored%202018_Time.PNG)
*Fig. 4: Refactored Nested 2018 Runtime

## Summary
### Advantages Vs Disadvantages Original
There are many advantages and disadvantages to both methods. An advantage would be that a nested for-loop is easier to implement. It is quite simple to follow, and you know what you are getting by following the iterations. This is an advantage over the arrays method because it is scalable at the cost of runtime. What I mean by scalable is we can know as little about our dataset and get the result we want. The disadvantage of the refactored code is that we need to know our variables. We had to create an additional index to keep count of and I felt like there can be more room for errors in terms of overflow because of the arrays. Another advantage would be that for a smaller dataset where runtime is minimal the nested for-loop method can be used and would save you debug time if you end up with more bugs by trying to refactor your code. 

### Advantages Vs Disadvantages Refactored
The advantage of the refactored code is that the runtime is only a multiple (N*O(n)) and not an exponential. When reading many more thousand lines of code or having a larger dataset (1GB+) it will take an exceptionally long time to calculate. The disadvantage to all this is that you must understand your dataset a bit more. You must know what you want from your data set and it cannot be as robust and modular as you want. In other words, you must write that code for a specific case like for our challenge, 12 green energy stocks. By creating multiple arrays, you are prone to more overflow errors and scalability by having to Redim the arrays. 
