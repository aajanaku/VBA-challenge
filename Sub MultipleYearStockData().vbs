Sub MultipleYearStockData()
    Dim ticker As String
    Dim nextrow As String
    Dim yearlyChange As Double
    Dim openValue As Double
    Dim closeValue As Double
    Dim i As Long
    Dim lastRow As Long
    Dim sheet As Worksheet
    Dim GreatestIncreaser As String 'ticker
    Dim GreatestValue As Double 'percent increase
    'Dim k As Long
    
    For Each sheet In ActiveWorkbook.Worksheets
        sheet.Cells(1, 9).Value = "ticker"
        sheet.Cells(1, 10).Value = "yearly change"
        sheet.Cells(1, 11).Value = "percent change"
        sheet.Cells(1, 12).Value = "total stock volume"
        sheet.Cells(1, 16).Value = "ticker"
        sheet.Cells(1, 17).Value = "value"
            
        'GreatestValue = Application.WorksheetFunction.Max(sheet.Range("K2:K3001")) 'This doesn't work
        
        'For k = 2 To 3001
            'If sheet.Cells(k, 17).Value = GreatestValue Then
                'GreatestIncreaser = sheet.Cells(k, 16).Value
                'Exit For
            'End If
        'Next k
        
        Dim totalValue As Double
        totalValue = 0
        
        openValue = sheet.Cells(2, 3).Value
        lastRow = sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
        
        Dim tickerRow As Long
        tickerRow = 2 'Assign ticker name
        
        ticker = sheet.Cells(tickerRow, 1).Value 'First ticker value
        
        For i = 2 To lastRow
            nextrow = sheet.Cells(i + 1, 1).Value
            
            If nextrow <> ticker Then ' Check if ticker has changed
                closeValue = sheet.Cells(i, 6).Value
                yearlyChange = closeValue - openValue
                
                sheet.Cells(tickerRow, 9).Value = ticker 'Assign ticker name
                sheet.Cells(tickerRow, 10).Value = yearlyChange
                If yearlyChange < 0 Then
                sheet.Cells(tickerRow, 10).Interior.Color = RGB(255, 0, 0)
                Else
                sheet.Cells(tickerRow, 10).Interior.Color = RGB(0, 255, 0)
                End If
                sheet.Cells(tickerRow, 11).Value = yearlyChange / openValue
                sheet.Cells(tickerRow, 12).Value = totalSV + sheet.Cells(i, 7).Value ' Add TSV
                
                tickerRow = tickerRow + 1 ' next row and ticker
                ticker = nextrow 'Ticker value for next group
                openValue = sheet.Cells(i + 1, 3).Value 'New ticker value
                totalVolume = 0 ' Reset TSV
            Else
                totalVolume = totalValue + sheet.Cells(i, 7).Value 'Total TSV
            End If
        Next i
    Next sheet
End Sub

