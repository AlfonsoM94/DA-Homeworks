Sub VolAdder()

'Identifies how many values are per sheet

 With ActiveSheet
    lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    MsgBox lastRow
End With

'Takes all the tickers
Dim CCTicker As Integer

Cells(2, 9).Value = Cells(2, 1).Value
CCTicker = 3

For i = 2 To lastRow
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Cells(CCTicker, 9).Value = Cells(i + 1, 1).Value
        CCTicker = CCTicker + 1
    End If

Next i

'Sums all the ticker's Amounts Charged

'Makes sure re-running the code won't duplicate or triplicate or so on and so forth values
For i = 2 To lastRow
    Cells(i, 9) = 0
Next i


For ii = 2 To lastRow
    For iVol = 2 To lastRow
        If Cells(ii, 1) = Cells(iVol, 9) Then
            Cells(iVol, 10).Value = Cells(iVol, 10).Value + Cells(ii, 7).Value
        End If
    Next iVol
    Next ii
    
End Sub