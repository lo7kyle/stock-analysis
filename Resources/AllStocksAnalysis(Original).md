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

