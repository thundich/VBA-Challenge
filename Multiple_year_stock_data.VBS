Sub Ticker()

'  Define Object Set
    Dim ws As Worksheet

'  Define last record of each worksheet
    Dim LastRecord As Double
    Dim LastSummaryRecord As Double
    Dim SaveTicker As String
    Dim SaveTickerOpen As Double
    Dim SaveTickerClose As Double
    Dim TickerSummaryRow As Integer
    Dim TickerVolume As Double
    Dim TickerChange As Double
    Dim MaxPercent As Double
    Dim MinPercent As Double
    Dim HighestVolume As Double
    Dim count As Integer

'  Set up loop to read all worksheets in the Workbook
    For Each ws In ActiveWorkbook.Worksheets


    '  Don't forget to activate each worksheet
        ws.Activate


    '  Get last record of each worksheet
        LastRecord = ws.Range("A1").End(xlDown).Row
     
    '  Write out column Headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

    '  For each worksheet, loop through each row
        TickerSymbolSummaryRow = 2
        For i = 2 To LastRecord
        
         ' If the first row of the loop, set default values
            If i = 2 Then
                SaveTicker = Cells(i, 1)
                SaveTickerOpen = Cells(i, 3)
                TickerVolume = 0
            End If


         ' Check if we are on the same ticker symbol
            If SaveTicker <> Cells(i, 1) Or i = LastRecord Then
                Cells(TickerSymbolSummaryRow, 9) = SaveTicker
                
             '  Set the ticker volume summary
                Cells(TickerSymbolSummaryRow, 12) = TickerVolume
                
             '  Calculate Yearly Change
                Cells(TickerSymbolSummaryRow, 10) = TickerClose - SaveTickerOpen
                If Cells(i, 10).Value > 0 Then
                    Cells(i, 10).Interior.ColorIndex = 10
                End If
                
                If Cells(i, 10).Value < 0 Then
                    Cells(i, 10).Interior.ColorIndex = 25
                End If
                
             '  Divisor cannot be 0 (divide by zero error will occur)
                If SaveTickerOpen <> 0 Then
                    Cells(TickerSymbolSummaryRow, 11) = Cells(TickerSymbolSummaryRow, 10) / SaveTickerOpen
                Else
                    Cells(TickerSymbolSummaryRow, 11) = 0
                End If
                Cells(TickerSymbolSummaryRow, 11).NumberFormat = "0.00%"
        
             '  Increment to the next Summary Row
                TickerSymbolSummaryRow = TickerSymbolSummaryRow + 1
                TickerVolume = 0

             '  Save Current TickerSymbol
                SaveTicker = Cells(i, 1)
                SaveTickerOpen = Cells(i, 3)
            End If
           
           
            If i <> LastRecord Then
                TickerVolume = TickerVolume + Cells(i, 7)
                TickerClose = Cells(i, 6)
            End If

        Next i
    
     '  Set Greatest % Increase, % Decrease, and Total Volume
        Cells(2, 14).Value = "Greatest % Increase"
        Cells(3, 14).Value = "Greatest % Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
    
     '  Get last record of each worksheet
        LastSummaryRecord = ws.Range("I1").End(xlDown).Row

    
    '   Look through the summary table to find the values and the associated ticker
        MinPercent = WorksheetFunction.Min(Range("K2:K" & LastSummaryRecord))
    
        HighestVolume = WorksheetFunction.Max(Range("L2:L" & LastSummaryRecord))
        
        MaxPercent = WorksheetFunction.Max(Range("K2:K" & LastSummaryRecord))
        
        For i = 2 To LastSummaryRecord
    
            If Cells(i, 11) = MaxPercent Then
                Cells(2, 15).Value = Cells(i, 9).Value
                Cells(2, 16).Value = MaxPercent
                Cells(2, 16).NumberFormat = "0.00%"
            End If
            If Cells(i, 11) = MinPercent Then
                Cells(3, 15).Value = Cells(i, 9).Value
                Cells(3, 16).Value = MinPercent
                Cells(3, 16).NumberFormat = "0.00%"
            End If
            If Cells(i, 12) = HighestVolume Then
                Cells(4, 15).Value = Cells(i, 9).Value
                Cells(4, 16).Value = HighestVolume
            
            End If
            
        Next i

    Next ws


End Sub

