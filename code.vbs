Sub contador()
For Each ws In Worksheets
Dim LastRow, MaxVolume, StockVolume As Long
Dim OldestDate, NewestDate, OpenValue, CloseValue, MaxValue, MinValue As Double
Dim LastColumnTicker, i, t, k As Long
Dim Title1, Title2, Title3, Title4 As String
Title1 = "Ticker"
Title2 = "Yearly Change"
Title3 = "Percentage change"
Title4 = "Total Stock Volume"
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Cells(1, 9).Value = Title1
Cells(1, 10).Value = Title2
Cells(1, 11).Value = Title3
Cells(1, 12).Value = Title4
For i = 2 To LastRow - 1
    If Cells(i, 1) <> Cells(i - 1, 1) Then
        LastColumnTicker = Cells(Rows.Count, 9).End(xlUp).Row
        Cells(LastColumnTicker + 1, 9).Value = Cells(i, 1).Value
    End If
Next i
OldestDate = Cells(2, 2).Value
For i = 2 To LastRow
    If Cells(i, 2) < OldestDate Then
        OldestDate = Cells(i, 2)
    End If
Next i
NewestDate = Cells(2, 2).Value
For i = 2 To LastRow
    If Cells(i, 2) >= NewestDate Then
        NewestDate = Cells(i, 2)
    End If
Next i
LastColumnTicker = Cells(Rows.Count, 9).End(xlUp).Row
For k = 2 To LastRow
    For t = 2 To LastColumnTicker
        If Cells(k, 1) = Cells(t, 9) And Cells(k, 2) = OldestDate Then
            OpenValue = Cells(k, 3)
        End If
        If Cells(k, 1) = Cells(t, 9) And Cells(k, 2) = NewestDate Then
            CloseValue = Cells(k, 6)
        Cells(t, 10) = CloseValue - OpenValue
        Cells(t, 11) = (CloseValue / OpenValue) - 1
        End If
    Next t
Next k
For k = 2 To LastColumnTicker
    StockVolume = 0
    For t = 2 To LastRow
        If Cells(t, 1) = Cells(k, 9) Then
        StockVolume = StockVolume + Cells(t, 7)
        Cells(k, 12).Value = StockVolume
        End If
    Next t
Next k
For k = 2 To LastColumnTicker
    If Cells(k, 10) <= 0 Then
        Cells(k, 10).Interior.ColorIndex = 3
   Else
        Cells(k, 10).Interior.ColorIndex = 43
    End If
Next k
Cells(2, 15).Value = "Greatest Increase"
MaxValue = Cells(2, 11).Value
For i = 2 To LastColumnTicker
    If Cells(i, 11) >= MaxValue Then
        MaxValue = Cells(i, 11).Value
        Cells(2, 17).Value = MaxValue
        Cells(2, 16).Value = Cells(i, 9).Value
    End If
Next i
Cells(3, 15).Value = "Greatest Decrease"
MinValue = Cells(2, 11).Value
For i = 2 To LastColumnTicker
    If Cells(i, 11) <= MinValue Then
        MinValue = Cells(i, 11).Value
        Cells(3, 17).Value = MinValue
        Cells(3, 16).Value = Cells(i, 9).Value
    End If
Next i
Cells(4, 15).Value = "Greatest Total Volume"
MaxVolume = Cells(2, 12).Value
For i = 2 To LastColumnTicker
    If Cells(i, 12) >= MaxVolume Then
        MaxVolume = Cells(i, 12).Value
        Cells(4, 17).Value = MaxVolume
        Cells(4, 16).Value = Cells(i, 9).Value
    End If
Next i
Next ws
End Sub