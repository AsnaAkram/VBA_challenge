Attribute VB_Name = "Module3"
Sub ticker_VBA():
    'creating variables
    Dim lastRow As Long
    Dim printCellRow As Integer
    Dim Ticker As String
    Dim sht As Worksheet
    Dim worksheetname As String
    
    
    Dim year_change As Variant
    Dim change As Variant
    Dim start As Variant
    Dim total_vol As Variant
    Dim percentage_ch As Variant
    
    For Each sht In Worksheets
        worksheetname = sht.Name
        MsgBox worksheetname
    
    total_vol = 0
    start = 2
    
    'nameing cells
    sht.Range("i1").Value = "Ticker"
    sht.Range("j1").Value = "Yearly Change"
    sht.Range("k1").Value = "Percentage Change"
    sht.Range("l1").Value = "Total Volume"
    
    printCellRow = 1
    Set sht = ActiveSheet
      lastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
    For Row = 2 To lastRow
        total_vol = total_vol + sht.Cells(Row, 7).Value
        
        If sht.Cells(Row, 1).Value <> sht.Cells(Row + 1, 1).Value Then
            sht.Cells(printCellRow + 1, 9).Value = sht.Cells(Row, 1).Value
        
    
            total_vol = total_vol + sht.Cells(Row, 7).Value
            sht.Cells(printCellRow + 1, 12).Value = total_vol
       
            change = sht.Cells(Row, 6) - sht.Cells(start, 3)
            percentage_ch = change / sht.Cells(start, 3)
            sht.Cells(printCellRow + 1, 10).Value = change
            sht.Cells(printCellRow + 1, 11).Value = percentage_ch
            
            percentage_ch = Format((change / start), "Percent")
                sht.Cells(printCellRow, 11) = Percent_Change
            
            If change > 0 Then
             sht.Range("J" & printCellRow + 1).Interior.ColorIndex = 4
             Else
             sht.Range("J" & printCellRow + 1).Interior.ColorIndex = 3
             End If
            
            printCellRow = printCellRow + 1
            total_vol = 0
            start = Row + 1
        
        End If
    Next Row
    Next sht
    
    
End Sub

