Attribute VB_Name = "Module1"
Sub analyzeStocks()

    'set static values for specific columsn
    Dim tickerColumn As Integer
    Dim openColumn As Integer
    Dim closeColumn As Integer
    Dim stockValueColumn As Integer
    tickerColumn = 1
    openColumn = 3
    closeColumn = 6
    stockValueColumn = 7

    'Loop through worksheets
    For Each ws In Worksheets
        
        'Label column headers
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change ($)"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
        'Choose first summary row
        Dim summaryRow As Integer
        summaryRow = 2
        
        'Determine last row for worksheet
        Dim lastRow As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim ticker As String
        Dim yearlyOpen As Double
        Dim yearlyClose As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim totalStockVolume As LongLong
        totalStockVolume = 0
        Debug.Print (ActiveSheet.Name)
        
        For i = 2 To lastRow
            
            'set open price for first ticker on each ws
            If i = 2 Then
                yearlyOpen = ws.Cells(i, openColumn)
            
            'check if ticker symbol is the different
            ElseIf ws.Cells(i + 1, tickerColumn).Value <> ws.Cells(i, tickerColumn).Value Then
                
                'store the ticker symbol
                ticker = ws.Cells(i, tickerColumn).Value
                
                'store final yearly change
                yearlyClose = ws.Cells(i, closeColumn)
                yearlyChange = yearlyClose - yearlyOpen
                
                 'determine yearly percent change. remove cases where dividing by 0 causes an error
                If yearlyOpen = 0 Then
                    percentChange = 0
                    
                Else
                    percentChange = yearlyChange / yearlyOpen
                    
                End If
                
                'store final stock volume
                totalStockVolume = totalStockVolume + ws.Cells(i, stockValueColumn).Value
                
                'write summary data
                ws.Range("I" & summaryRow).Value = ticker
                ws.Range("J" & summaryRow).Value = yearlyChange
                    
                    'format yearly change color, "conditional formatting"
                    If yearlyChange > 0 Then
                        ws.Range("J" & summaryRow).Interior.ColorIndex = 4
                    ElseIf yearlyChange < 0 Then
                        ws.Range("J" & summaryRow).Interior.ColorIndex = 3
                    End If
                    
                ws.Range("K" & summaryRow).Value = percentChange
                ws.Range("K2:K" & summaryRow).NumberFormat = "0.00%"
                ws.Range("L" & summaryRow).Value = totalStockVolume
                
                'Leave for loop if finished
                If i = lastRow Then
                    Debug.Print ("exit")
                    Exit For
                End If
                
                'reset stock volume
                totalStockVolume = 0
                'set yearly open price
                yearlyOpen = ws.Cells(i + 1, openColumn).Value

                'move summary row
                summaryRow = summaryRow + 1

            Else
                'ticker symbol is the same
                totalStockVolume = totalStockVolume + ws.Cells(i, stockValueColumn).Value
            
            End If
        
        Next i
        
        'label bonus statistics table
        ws.Range("N2") = "Greatest % Increase"
        ws.Range("N3") = "Greatest % Decrease"
        ws.Range("N4") = "Greatest Total Volume"
        ws.Range("O1") = "Ticker"
        ws.Range("P1") = "Value"
        
        'set generated data as a new table
        Dim summaryTable As Range
        Dim endRowSumTable As Long
        endRowSumTable = ws.Cells(Rows.Count, 9).End(xlUp).Row
        Set summaryTable = ws.Range("I2:K" & endRowSumTable)
        
        'find bonus statistics
        Dim greatestPercentInc As Double
        Dim greatestPercentDec As Double
        Dim greatestTotalVol As LongLong
        greatestPercentInc = WorksheetFunction.Max(summaryTable.Columns(3))
        greatestPercentDec = WorksheetFunction.Min(summaryTable.Columns(3))
        greatestTotalVol = WorksheetFunction.Max(summaryTable.Columns(4))
        
        'print out bonus statistics. Use Index/Match to find ticker symbols
        ws.Range("O2").Value = WorksheetFunction.Index(summaryTable.Columns(1), WorksheetFunction.Match(greatestPercentInc, summaryTable.Columns(3), 0))
        ws.Range("O3").Value = WorksheetFunction.Index(summaryTable.Columns(1), WorksheetFunction.Match(greatestPercentDec, summaryTable.Columns(3), 0))
        ws.Range("O4") = WorksheetFunction.Index(summaryTable.Columns(1), WorksheetFunction.Match(greatestTotalVol, summaryTable.Columns(4), 0))
        ws.Range("P2") = FormatPercent(greatestPercentInc, 2)
        ws.Range("P3") = FormatPercent(greatestPercentDec, 2)
        ws.Range("P4") = greatestTotalVol
        
    Next ws

End Sub


