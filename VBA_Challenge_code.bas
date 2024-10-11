Attribute VB_Name = "Module1"
Sub Loop_Multi_Stock_Data()
    ' Create a script that loops through all the stocks for each quarter and output
    ' New outputs for quarterly sum
    
    Dim WS As Worksheet
    Dim LastRow As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim QuarterlyChanges As Double
    Dim TickerSymbol As String
    Dim PercentageChange As Double
    Dim Volume As Double
    Dim Row As Long
    Dim r As Long
    Dim j As Long
    Dim O As Long
    Dim YCLastRow As Long
    
    For Each WS In ActiveWorkbook.Worksheets
        WS.Activate
        ' Determine the Last Row
        LastRow = WS.Cells(Rows.Count, "A").End(xlUp).Row
        
        ' Name Cell Values
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Quarterly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        ' Initialize Variables
        Volume = 0
        Row = 2
        
        ' Set initial Opening Price
        OpeningPrice = Cells(2, "C").Value
        
        ' Loop through all ticker symbols
        For r = 2 To LastRow
            If Cells(r + 1, "A").Value <> Cells(r, "A").Value Then
                ' Set Ticker Symbol
                TickerSymbol = Cells(r, "A").Value
                Cells(Row, "I").Value = TickerSymbol
                
                ' Set Closing Price
                ClosingPrice = Cells(r, "F").Value
                
                ' Set Quarterly Change
                QuarterlyChanges = ClosingPrice - OpeningPrice
                Cells(Row, "J").Value = QuarterlyChanges
                
                ' Set Percentage Change
                If (OpeningPrice = 0 And ClosingPrice = 0) Then
                    PercentageChange = 0
                ElseIf (OpeningPrice = 0 And ClosingPrice <> 0) Then
                    PercentageChange = 1
                Else
                    PercentageChange = QuarterlyChanges / OpeningPrice
                End If
                Cells(Row, "K").Value = PercentageChange
                Cells(Row, "K").NumberFormat = "0.00%"
                
                ' Add Total Volume
                Volume = Volume + Cells(r, "G").Value
                Cells(Row, "L").Value = Volume
                
                ' Move to the next summary table row
                Row = Row + 1
                
                ' Reset the Opening Price
                OpeningPrice = Cells(r + 1, "C").Value ' Make sure to reference the next row
                
                ' Reset the Volume Total
                Volume = 0
            Else
                ' If cells are the same ticker
                Volume = Volume + Cells(r, "G").Value
            End If
        Next r
        
        ' Determine the Last Row of Yearly Change per WS
        YCLastRow = WS.Cells(Rows.Count, "I").End(xlUp).Row
        
       ' Set the Cell Colors for Column J
        For j = 2 To YCLastRow
        If Cells(j, "J").Value > 0 Then
        ' Green for positive values
        Cells(j, "J").Interior.ColorIndex = 10
        ElseIf Cells(j, "J").Value < 0 Then
        ' Red for negative values
        Cells(j, "J").Interior.ColorIndex = 3
        Else
        ' No color for zero
        Cells(j, "J").Interior.ColorIndex = xlNone
        End If
        Next j
        
        ' Add functionality to your script to return the stock with
        ' "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
        Cells(2, "N").Value = "Greatest % Increase"
        Cells(3, "N").Value = "Greatest % Decrease"
        Cells(4, "N").Value = "Greatest Total Volume"
        Cells(1, "O").Value = "Ticker"
        Cells(1, "P").Value = "Value"
        
        ' Look through each row to find the greatest value and its associated ticker
        For O = 2 To YCLastRow
            If Cells(O, "K").Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, "O").Value = Cells(O, "A").Value
                Cells(2, "P").Value = Cells(O, "K").Value
                Cells(2, "P").NumberFormat = "0.00%"
            ElseIf Cells(O, "K").Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, "O").Value = Cells(O, "A").Value
                Cells(3, "P").Value = Cells(O, "K").Value
                Cells(3, "P").NumberFormat = "0.00%"
            ElseIf Cells(O, "L").Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, "O").Value = Cells(O, "A").Value
                Cells(4, "P").Value = Cells(O, "L").Value
            End If
        Next O
    Next WS
End Sub

