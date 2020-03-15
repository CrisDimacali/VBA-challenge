Sub Alphabetical_testing()
For Each ws In Worksheets
    Dim worksheetName As String
    Dim lastrow As Long
    Dim VolStock As Double
    Dim Vol As Long
    Dim tableRow As Long
    Dim i As Long
    Dim ticker As String
    Dim yearopen As Double
    Dim yearclose As Double
    Dim greatestIncrease As Double
    yearopen = 0
    VolStock = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    tableRow = 2
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly_Change"
    ws.Cells(1, 11).Value = "Percent_Change"
    ws.Cells(1, 12).Value = "Total_Stock_Volume"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
 
    For i = 2 To lastrow
        VolStock = VolStock + ws.Cells(i, 7).Value
        If yearopen = 0 Then
           yearopen = ws.Cells(i, 3).Value
      End If
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i - 1, 1) = ws.Cells(i, 1) Then
            yearclose = ws.Cells(i, 6).Value
            yearlychange = yearclose - yearopen
            If yearopen = 0 Then
                yearlypercent = "N/A"
            Else
                yearlypercent = yearlychange / yearopen
            End If
            
            Range("K" & tableRow).NumberFormat = "0.00%"
            ws.Cells(tableRow, 9).Value = ws.Cells(i, 1)
            ws.Cells(tableRow, 10).Value = yearlychange
            ws.Cells(tableRow, 11).Value = yearlypercent
            ws.Cells(tableRow, 12).Value = VolStock
            
            If yearlychange > 0 Then
               ws.Cells(tableRow, 21) = yearlychange
               ws.Cells(tableRow, 10).Interior.ColorIndex = 4
            ElseIf yearlychange <= 0 Then
                ws.Cells(tableRow, 22) = yearlychange
                ws.Cells(tableRow, 10).Interior.ColorIndex = 3
            End If
            
            If yearlypercent <> "N/A" And yearlypercent >= 0 And (tableRow = 2 Or yearlypercent > ws.Cells(2, 17).Value) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 1)
                ws.Cells(2, 17).Value = yearlypercent
            End If
           
            If yearlypercent <> "N/A" And yearlypercent < 0 And (tableRow = 2 Or yearlypercent < ws.Cells(3, 17).Value) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 1)
                ws.Cells(3, 17).Value = yearlypercent
            End If
           
            If tableRow = 2 Or VolStock > ws.Cells(4, 17).Value Then
                ws.Cells(4, 16).Value = ws.Cells(i, 1)
                ws.Cells(4, 17).Value = VolStock
            End If
           
            tableRow = tableRow + 1
            VolStock = 0
            yearopen = 0
        End If
Next i
   ws.Cells(2, 17).NumberFormat = "0.00%"
   ws.Cells(3, 17).NumberFormat = "0.00%"
Next ws

End Sub


