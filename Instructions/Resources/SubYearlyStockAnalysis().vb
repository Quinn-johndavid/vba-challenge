Sub YearlyStockAnalysis()


'Variable declare

Dim ws As Worksheet, rowCount As Long, firstCapture As Long, count As Long, errorCount As Long, lastRow As Long, yearValChange As Double, percentageChange As Double, stockVolTotal As Double, tickerNameArray(), tickerValArray()

'worksheet iteration

    For Each ws In ThisWorkbook.Worksheets
        
        Worksheets(ws.Name).Activate
    
        'Sorting arguements, using built in functions because i thought they were neat and easier to use!!!
        Range("A2").End(xlDown).End(xlToRight).Sort [A2], xlAscending, Header:=xlYes
        
        'setting Last Row
        lastRow = Cells(Rows.count, 1).End(xlUp).Row
        
        'Summary labels location and names
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'setting variables = to 0
        'Fixes my div by 0 error i got in the last hw assignment
        count = 0
        errorCount = 0
        stockVolTot = 0
        
        'Start with Row2, when the loop finds a new ticker that isnt in its array, firstCapture adds it
        firstCapture = 2
        For rowCount = 2 To lastRow
            
            If Cells(rowCount, 1).Value <> Cells(rowCount + 1, 1).Value Then
            
               'Check if there is a change in value of Ticker'
                count = count + 1
        
                'Calc yearly changes
                yearValChange = Cells(rowCount, 6).Value - Cells(firstCapture, 3).Value
        
                'calc percentage change
                If Cells(firstCapture, 3).Value = 0 Then
                    percentageChange = 0
                    If errorCount = 0 Then
                        'Labels are added on first Error if error
                        Cells(6, 16).Value = "Tickers with Errors(divide by 0 err)"
                        Cells(6, 16).Interior.ColorIndex = 6
                    End If
                    Cells(7 + errorCount, 16).Value = Cells(rowCount, 1).Value
                    Cells(7 + errorCount, 16).Interior.ColorIndex = 6
                    errorCount = errorCount + 1
                Else
                    percentageChange = yearValChange / Cells(firstCapture, 3).Value
                End If
        
                'Calculate stockVolTot
                stockVolTot = stockVolTot + Cells(rowCount, 7).Value
        
                
                Cells(count + 1, 9).Value = Cells(rowCount, 1).Value
                Cells(count + 1, 10).Value = yearValChange
                Cells(count + 1, 11).Value = percentageChange
                Cells(count + 1, 12).Value = stockVolTot
        
                'Conditional Formatting for yearvalchange
                If yearValChange > 0 Then
                    Cells(count + 1, 10).Interior.ColorIndex = 4 'Green
                Else
                    Cells(count + 1, 10).Interior.ColorIndex = 3 'Red
                End If
        
                
                If count = 1 Then
                'Initialize the arrays
                'then pull the data that meets the if statements arguments, goes in order of the list below
                'Greatest % increase tracking
                'Greatest % decrease tracking
                'Greatest stock volume tracking

                    tickerNameArray = Array(Cells(rowCount, 1), Cells(rowCount, 1), Cells(rowCount, 1))
                    tickerValueArray = Array(percentageChange, percentageChange, stockVolTot)
                Else
        
                        If percentageChange > tickerValueArray(0) Then
                            tickerValueArray(0) = percentageChange
                            tickerNameArray(0) = Cells(rowCount, 1)
                        End If
                
                        If percentageChange < tickerValueArray(1) Then
                            tickerValueArray(1) = percentageChange
                            tickerNameArray(1) = Cells(rowCount, 1)
                        End If
                
                        If stockVolTot > tickerValueArray(2) Then
                            tickerValueArray(2) = stockVolTot
                            tickerNameArray(2) = Cells(rowCount, 1)
                        End If
            
                End If
                
                'Updates and moves on
                firstCapture = rowCount + 1
                stockVolTot = 0
                
            Else
                'update the stockVolTot
                stockVolTot = stockVolTot + Cells(rowCount, 7).Value
            End If
            
        Next rowCount
        
        
        
        'Label Filling then filling in the
        'Results
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        
        Cells(2, 16).Value = tickerNameArray(0)
        Cells(3, 16).Value = tickerNameArray(1)
        Cells(4, 16).Value = tickerNameArray(2)
        
        Cells(2, 17).Value = tickerValueArray(0)
        Cells(3, 17).Value = tickerValueArray(1)
        Cells(4, 17).Value = tickerValueArray(2)
        
        
        'Number and  Column width formatting, would use a method to call this but dont know how to do that in vba, 
        'also would like to know how to clean worksheets with executed macros so i dont have to go without saving my changes to see if my code works
        Columns(10).NumberFormat = "0.00"
        Columns(11).NumberFormat = "0.00%"
        Range(Cells(2, 17), Cells(3, 17)).NumberFormat = "0.00%"
        Cells(4, 17).NumberFormat = "0.0000E+00"
        
        Columns("H:H").ColumnWidth = 44
        Columns("I:I").ColumnWidth = 14
        Columns("J:J").ColumnWidth = 14
        Columns("K:K").ColumnWidth = 14
        Columns("L:L").ColumnWidth = 20
        Columns("M:M").ColumnWidth = 22
        Columns("N:N").ColumnWidth = 22
        Columns("O:O").ColumnWidth = 26
        Columns("P:P").ColumnWidth = 14
        Columns("Q:Q").ColumnWidth = 20
        Range("I1:Q1").Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
    Next

    

End Sub



