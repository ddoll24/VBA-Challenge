Attribute VB_Name = "Module1"
Sub StockData()
    ' Loop through all sheets
    
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        ' Determine the Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Add category for summary
        
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        'Create variables to hold values
        
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim TickerName As String
        Dim PercentChange As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        'Set the initial open price
        OpenPrice = Cells(2, Column + 2).Value
        
         ' Loop through all of the ticker symbols
        
        For i = 2 To LastRow
        
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                
                ' Set the ticker name
                
                TickerName = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = TickerName
                
                ' Set the close price
                
                ClosePrice = Cells(i, Column + 5).Value
                
                ' Add the yearly change
                
                YearlyChange = ClosePrice - OpenPrice
                Cells(Row, Column + 9).Value = YearlyChange
                
                ' Add percent change
                
                If (OpenPrice = 0 And ClosePrice = 0) Then
                    PercentChange = 0
                ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
                    PercentChange = 1
                Else
                    PercentChange = YearlyChange / OpenPrice
                    Cells(Row, Column + 10).Value = PercentChange
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                
                ' Add total volumn
                
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                
                ' Add one to the summary table row
                
                Row = Row + 1
                
                ' Reset the open price
                
                OpenPrice = Cells(i + 1, Column + 2)
                
                ' Reset the volumn total
                
                Volume = 0
                
            'If cells are the same ticker
            
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
        Next i
        
        ' Determine the last row of yearly change per WS
        
        YCLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        
        ' Set the cell colors using conditional formatting and color index
        
        For j = 2 To YCLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 4
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        ' Set the greatest % increase, greatest % decrease, and total volume
        
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
        
        ' Look through each row to find the greatest value and the corresponding ticker
        
        For Z = 2 To YCLastRow
            If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
            End If
        Next Z
        
    Next WS
        
End Sub
