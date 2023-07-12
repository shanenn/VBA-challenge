Attribute VB_Name = "Module1"
Sub iterate()
For Each Page In ActiveWorkbook.Worksheets
    Page.Cells(1, 9) = "Ticker"
    Page.Cells(1, 10) = "Yearly Change"
    Page.Cells(1, 11) = "Percent Change"
    Page.Cells(1, 12) = "Total Stock Volume"
    tick = Page.Cells(2, 1)
    year_start = Page.Cells(2, 3)
    vol = Page.Cells(2, 7)
    summind = 2
    For i = 3 To Page.Range("A2").End(xlDown).Row
        If tick = Page.Cells(i, 1) Then
            vol = vol + Page.Cells(i, 7)
            year_end = Page.Cells(i, 6)
        Else
            Page.Cells(summind, 9) = tick
            Page.Cells(summind, 10) = year_end - year_start
            Page.Cells(summind, 10).NumberFormat = "0.00"
            If Page.Cells(summind, 10) > 0 Then
                Page.Cells(summind, 10).Interior.ColorIndex = 4
            ElseIf Page.Cells(summind, 10) < 0 Then
                Page.Cells(summind, 10).Interior.ColorIndex = 3
            End If
            Page.Cells(summind, 11) = Page.Cells(summind, 10) / year_start
            Page.Cells(summind, 11).NumberFormat = "0.00%"
            Page.Cells(summind, 12) = vol
            tick = Page.Cells(i, 1)
            year_start = Page.Cells(i, 3)
            vol = Page.Cells(i, 7)
            summind = summind + 1
        
        End If
    Next i
    Page.Cells(summind, 9) = tick
    Page.Cells(summind, 10) = year_end - year_start
    Page.Cells(summind, 10).NumberFormat = "0.00"
    If Page.Cells(summind, 10) > 0 Then
        Page.Cells(summind, 10).Interior.ColorIndex = 4
    ElseIf Page.Cells(summind, 10) < 0 Then
        Page.Cells(summind, 10).Interior.ColorIndex = 3
    End If
    Page.Cells(summind, 11) = Page.Cells(summind, 10) / year_start
    Page.Cells(summind, 11).NumberFormat = "0.00%"
    Page.Cells(summind, 12) = vol
    
    dec = 0
    inc = 0
    tot = 0
    For i = 2 To Page.Range("I2").End(xlDown).Row
        If Page.Cells(i, 11) > inc Then
            inc = Page.Cells(i, 11)
            inctick = Page.Cells(i, 9)
        End If
        If Page.Cells(i, 11) < dec Then
            dec = Page.Cells(i, 11)
            dectick = Page.Cells(i, 9)
        End If
        If Page.Cells(i, 12) > tot Then
            tot = Page.Cells(i, 12)
            tottick = Page.Cells(i, 9)
        End If
    Next i
    
    Page.Cells(2, 15) = "Greatest % Increase"
    Page.Cells(3, 15) = "Greatest % Decrease"
    Page.Cells(4, 15) = "Greatest Total Volume"
    Page.Cells(1, 16) = "Ticker"
    Page.Cells(1, 17) = "Value"
    Page.Cells(2, 16) = inctick
    Page.Cells(2, 17) = inc
    Page.Cells(2, 17).NumberFormat = "0.00%"
    Page.Cells(3, 16) = dectick
    Page.Cells(3, 17) = dec
    Page.Cells(3, 17).NumberFormat = "0.00%"
    Page.Cells(4, 16) = tottick
    Page.Cells(4, 17) = tot
    
            

Next Page
End Sub

