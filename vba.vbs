Sub VBAStocks():

    For Each WS In Worksheets

        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"
    
        WS.Cells(2, 14).Value = "Greatest % Increase"
        WS.Cells(3, 14).Value = "Greatest % Decrease"
        WS.Cells(4, 14).Value = "Greatest Total Volume"
    
        WS.Cells(1, 15).Value = "Ticker"
        WS.Cells(1, 16).Value = "Value"
        
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim TickerName As String
        Dim PercentChange As Double
        Dim Volume As Double
        Volume = 0
        Dim R As Long
        R = 2
        Dim C As Integer
        C = 1
        Dim i As Long
        'Dim StartRow As Long
        'StartRow = 2
        OpenPrice = WS.Cells(2, 3).Value
        Dim z As Long
         
        For i = 2 To LastRow
            Volume = Volume + WS.Cells(i, C + 6).Value
            If WS.Cells(i + 1, C).Value <> WS.Cells(i, C).Value Then
                TickerName = WS.Cells(i, C).Value
                WS.Cells(R, C + 8).Value = TickerName
                
                'OpenPrice = WS.Cells(StartRow, C + 2).Value
                ClosePrice = WS.Cells(i, C + 5).Value

                YearlyChange = ClosePrice - OpenPrice
                WS.Cells(R, C + 9).Value = YearlyChange

                If (OpenPrice = 0) Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / OpenPrice
                
                WS.Cells(R, C + 10).Value = PercentChange
                WS.Cells(R, C + 10).NumberFormat = "0.00%"
                WS.Cells(R, C + 11).Value = Volume
                
                R = R + 1
                'StartRow = i + 1
                OpenPrice = WS.Cells(i + 1, 3)
                Volume = 0
            
                End If
            End If
        Next i
        
        YearlyChangeLastRow = WS.Cells(Rows.Count, C + 8).End(xlUp).Row
        
        For j = 2 To YearlyChangeLastRow
            If (WS.Cells(j, C + 9).Value > 0 Or WS.Cells(j, C + 9).Value = 0) Then
                WS.Cells(j, C + 9).Interior.ColorIndex = 43
            ElseIf WS.Cells(j, C + 9).Value < 0 Then
                WS.Cells(j, C + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        For z = 2 To YearlyChangeLastRow
            If WS.Cells(z, C + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YearlyChangeLastRow)) Then
                WS.Cells(2, C + 14).Value = WS.Cells(z, C + 8).Value
                WS.Cells(2, C + 15).Value = WS.Cells(z, C + 10).Value
                WS.Cells(2, C + 15).NumberFormat = "0.00%"
            End If
            If WS.Cells(z, C + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YearlyChangeLastRow)) Then
                WS.Cells(3, C + 14).Value = WS.Cells(z, C + 8).Value
                WS.Cells(3, C + 15).Value = WS.Cells(z, C + 10).Value
                WS.Cells(3, C + 15).NumberFormat = "0.00%"
            End If
            If WS.Cells(z, C + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YearlyChangeLastRow)) Then
                WS.Cells(4, C + 14).Value = WS.Cells(z, C + 8).Value
                WS.Cells(4, C + 15).Value = WS.Cells(z, C + 11).Value
            End If
        Next z
    
    Next WS
    
End Sub
