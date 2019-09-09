Sub Stock_Analysis()
       
    'This is a loop to read and process the data in the worksheets in the workbook
    'Set a variable to read data from the worksheets
    Dim work As Worksheet
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
        

    'Begin loop for worksheets
    For Each work In Worksheets

        'Create and label columns Ticker,Yearly change, Percent Change and Total Stock Volume
        'for a summary table to hold data
        work.Cells(1, 9).Value = "Ticker"
        work.Cells(1, 10).Value = "Yearly Change"
        work.Cells(1, 11).Value = "Percent Change"
        work.Cells(1, 12).Value = "Total Stock Volume"

        'Set ticker symbol variable
        Dim tickersym As String

        'Set total volume of stock traded variable
        Dim totalvol As Double
        totalvol = 0

        'Variable to track location for each ticker symbol in the summary table
        Dim rowcount As Long
        rowcount = 2

        'Declare year open price variable
        Dim yearopen As Double
        yearopen = 0

        'Declare year close price variable
        Dim yearclose As Double
        yearclose = 0
        
        'Declare the change in price for the year variable
        Dim yearchange As Double
        yearchange = 0

        'Declare the % change in price for the year variable
        Dim perchg As Double
        perchg = 0

        'Declare total rows to loop through variable
        Dim lastrow As Long
        lastrow = work.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop iterations for ticker symbols
        For tick = 2 To lastrow
            
            'IF Condition to get year open price
            If work.Cells(tick, 1).Value <> work.Cells(tick - 1, 1).Value Then

                yearopen = work.Cells(tick, 3).Value

            End If

            'Sum up the volume for each row to calculate the total stock volume for the year
            totalvol = totalvol + work.Cells(tick, 7)

            'IF Condition to check if the ticker symbol has changed
            If work.Cells(tick, 1).Value <> work.Cells(tick + 1, 1).Value Then

                'Transfer ticker symbol to summary table
                work.Cells(rowcount, 9).Value = work.Cells(tick, 1).Value

                'Transfer total stock volume to the summary table
                work.Cells(rowcount, 12).Value = totalvol

                'Read year closing price
                yearclose = work.Cells(tick, 6).Value

                'Calculate the price change for the year and transfer it to the summary table.
                yearchange = yearclose - yearopen
                work.Cells(rowcount, 10).Value = WorksheetFunction.Ceiling(yearchange, 0.000000005)
                                                                                                
                'IF Condition to format and highlight green for positive and red for negative change.
                If yearchange >= 0 Then
                    work.Cells(rowcount, 10).Interior.ColorIndex = 4
                Else
                    work.Cells(rowcount, 10).Interior.ColorIndex = 3
                End If
                
                
                'Calculate the % change for the year and transfer to the summary table as % format
                'IF condition to calculate % change
                'Check if beginning and ending values are zero. This will show no increase and formula won't work if we divide by 0
                'If a stock begins at a value of zero and increases, it will show an infinite % increase.
                'So we only need to look at actual price increase in $ and say "This Stock is New" instead of percent change.
                If yearopen = 0 And yearclose = 0 Then
                    perchg = 0
                    work.Cells(rowcount, 11).Value = perchg
                    work.Cells(rowcount, 11).NumberFormat = "0.00%"
                ElseIf yearopen = 0 Then
                    Dim perchg_NA As String
                    perchg_NA = "This stock is New"
                    work.Cells(rowcount, 11).Value = perchg
                Else
                    perchg = yearchange / yearopen
                    work.Cells(rowcount, 11).Value = perchg
                    work.Cells(rowcount, 11).NumberFormat = "0.00%"
                End If

                'Add 1 to rowcount to transfer it to the next empty row in the summary table
                rowcount = rowcount + 1

                'Reset total stock volume, year open price, year close price, year change, year percent change
                totalvol = 0
                yearopen = 0
                yearclose = 0
                yearchange = 0
                perchg = 0
                
            End If
        Next tick

        'Create table with headings for best/worst performance
        work.Cells(2, 15).Value = "Greatest % Increase"
        work.Cells(3, 15).Value = "Greatest % Decrease"
        work.Cells(4, 15).Value = "Greatest Total Volume"
        work.Cells(1, 16).Value = "Ticker"
        work.Cells(1, 17).Value = "Value"

        'Set lastrow to count the number of rows in the summary table
        lastrow = work.Cells(Rows.Count, 9).End(xlUp).Row

        'Set variables for best performer, worst performer, and stock with the most volume
        Dim beststock As String
        Dim bestvalue As Double

        'Set best performer equal to the first stock
        bestvalue = work.Cells(2, 11).Value

        Dim worststock As String
        Dim worstvalue As Double

        'Set worst performer equal to the first stock
        worstvalue = work.Cells(2, 11).Value

        Dim mostvolstock As String
        Dim mostvolvalue As Double

        'Set most volume equal to the first stock
        mostvolvalue = work.Cells(2, 12).Value

        'Loop to read through summary table
        For perf = 2 To lastrow

            'If condition to find best performer
            If work.Cells(perf, 11).Value > bestvalue Then
                bestvalue = work.Cells(perf, 11).Value
                beststock = work.Cells(perf, 9).Value
            End If

            'IF condition to find worst performer
            If work.Cells(perf, 11).Value < worstvalue Then
                worstvalue = work.Cells(perf, 11).Value
                worststock = work.Cells(perf, 9).Value
            End If

            'IF condition to find stock with the greatest volume traded
            If work.Cells(perf, 12).Value > mostvolvalue Then
                mostvolvalue = work.Cells(perf, 12).Value
                mostvolstock = work.Cells(perf, 9).Value
            End If

        Next perf

        'transfer best performer, worst performer, and stock with the most volume items to the performance table
        work.Cells(2, 16).Value = beststock
        work.Cells(2, 17).Value = bestvalue
        work.Cells(2, 17).NumberFormat = "0.00%"
        work.Cells(3, 16).Value = worststock
        work.Cells(3, 17).Value = worstvalue
        work.Cells(3, 17).NumberFormat = "0.00%"
        work.Cells(4, 16).Value = mostvolstock
        work.Cells(4, 17).Value = mostvolvalue

        'Autofit table columns
        work.Columns("I:L").EntireColumn.AutoFit
        work.Columns("O:Q").EntireColumn.AutoFit

    Next work
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub