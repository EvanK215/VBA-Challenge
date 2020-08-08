Attribute VB_Name = "Stocks"
   
Sub stockrecap()
Dim column As Integer
Dim ws As Worksheet
Dim FirstTickerOpen As Double
Dim yearpctchng As Double
Dim YearPriceChng As Double

'Loop Each Worksheet
For Each ws In Worksheets
ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Yr Price Chng"
ws.Range("L1").Value = "Yr Pct Chng"
ws.Range("M1").Value = "Total Vol"

column = 1
ws.Range("j2:q1000000").ClearContents
'counts the number of rows
LastRowData = ws.Cells(Rows.Count, 1).End(xlUp).Row

FirstTickerOpen = ws.Cells(2, 3)

' Loop through rows in the column
    For i = 2 To LastRowData
    'searches for when the value of the next cell is different than that of the current cell
        If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
            lastTickerClose = ws.Cells(i, 6).Value
            'find next rown in summary in summary table
            lastRowSummary = ws.Cells(i, 10).End(xlUp).Row
            ' add unique stock ticker to summary
            ws.Cells(lastRowSummary + 1, 10).Value = ws.Cells(i, column).Value
            'Calculate yearly price change
            YearPriceChng = lastTickerClose - FirstTickerOpen
            'display YearPriceChange
            ws.Cells(lastRowSummary + 1, 11).Value = YearPriceChng
            'calculate Persent yeral change
            If FirstTickerOpen = 0 Then
                yearpctchng = 0
            Else
                yearpctchng = YearPriceChng / FirstTickerOpen
            End If
            ws.Cells(lastRowSummary + 1, 12).Value = yearpctchng
            ws.Cells(lastRowSummary + 1, 12).NumberFormat = "0.00%"
                'format Pctchange cell positive or negative
                If yearpctchng >= 0 Then
                   ws.Cells(lastRowSummary + 1, 11).Interior.Color = vbGreen
                Else
                    ws.Cells(lastRowSummary + 1, 11).Interior.Color = vbRed
                End If
            ' Add to the Ticker Volume
              TickerTotal = TickerTotal + ws.Cells(i, 7).Value
              ' Print Ticker volu,me total
              ws.Cells(lastRowSummary + 1, 13).Value = TickerTotal
              ws.Cells(lastRowSummary + 1, 13).NumberFormat = "#,##0"
              ' Reset the Ticker Total
              TickerTotal = 0
              'Rest First open of next ticker
               FirstTickerOpen = ws.Cells(i + 1, 3).Value
        Else
           ' Add to the Ticker Volume
             TickerTotal = TickerTotal + ws.Cells(i, 7).Value
 
        End If
    Next i
    
'Challenge totals
'Find range of summary table
lastSummaryData = ws.Cells(Rows.Count, 12).End(xlUp).Row
'set labels for challenge totals
ws.Range("o2").Value = "Max Yr Pct Chng"
ws.Range("o3").Value = "Min Yr Pct Chng"
ws.Range("o4").Value = "Max Volume"
ws.Range("p1").Value = "Ticker"
ws.Range("q1").Value = "Amount"

    For j = 2 To lastSummaryData
        'Find Highest Yearly % increase
            If ws.Cells(j, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & lastSummaryData)) Then
                ws.Cells(2, 16).Value = ws.Cells(j, 10).Value
                ws.Cells(2, 17).Value = ws.Cells(j, 12).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
        'Find Highest Yearly % increase
            ElseIf ws.Cells(j, 12).Value = WorksheetFunction.Min(ws.Range("L2:L" & lastSummaryData)) Then
                ws.Cells(3, 16).Value = ws.Cells(j, 10).Value
                ws.Cells(3, 17).Value = ws.Cells(j, 12).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
             End If
        'Find Highest total yearly volume
            If ws.Cells(j, 13).Value = WorksheetFunction.Max(ws.Range("M2:M" & lastSummaryData)) Then
                ws.Cells(4, 16).Value = ws.Cells(j, 10).Value
                ws.Cells(4, 17).Value = ws.Cells(j, 13).Value
                ws.Cells(4, 17).NumberFormat = "#,##0"
            End If
    Next j
    ws.Columns("A:Q").AutoFit
    ws.Range("j1").Select
  
  Next ws
 
 
End Sub
