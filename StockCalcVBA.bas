Attribute VB_Name = "Module1"
Sub tickerCalcs():
     Dim firstOpenPrice As Double
     Dim lastClosePrice As Double
     Dim summaryIterator As Integer
     Dim openColumn As Integer
     Dim clsColumn As Integer
     Dim tickColumn As Integer
     Dim volColumn As Integer
     Dim biggestIncr As Double
     Dim biggestDecr As Double
     Dim mostVol As LongLong
     Dim totalVol As LongLong
     Dim biggestIncrT As String
     Dim biggestDecrT As String
     Dim mostVolT As String
     
     volColumn = 7
     tickColumn = 1
     openColumn = 3
     clsColumn = 6
     
     
     
   For Each ws In Worksheets
        'reset "perSheet" variables
        summaryIterator = 2
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'set "biggest things" to the 1st ticker/0 for now
        biggestIncrT = ws.Cells(2, tickColumn).Value
        biggestIncr = 0
        biggestDecrT = ws.Cells(2, tickColumn).Value
        biggestDecr = 0
        mostVolT = ws.Cells(2, tickColumn).Value
        mostVol = 0
   
      ' loop through each row in the sheet:
      For i = 2 To lastRow
          If curTicker = "" Then
             'check if curTicker is set, if it's not- set the curTicker, openPrice
              curTicker = ws.Cells(i, tickColumn).Value
              firstOpenPrice = ws.Cells(i, openColumn).Value
          End If
          currentRow = ws.Cells(i, tickColumn).Value
          nextRow = ws.Cells(i + 1, tickColumn).Value
   
          If currentRow = nextRow Then
            ' MsgBox ("samesame")
             totalVol = totalVol + ws.Cells(i, volColumn).Value
          Else 'next row starts different ticker
            ' MsgBox ("Different!")
             totalVol = totalVol + ws.Cells(i, volColumn).Value
             
             'Update summary table for current ticker:
             ws.Cells(summaryIterator, 9).Value = currentRow
             yearChange = ws.Cells(i, clsColumn).Value - firstOpenPrice '
             ws.Cells(summaryIterator, 10).Value = yearChange
             If yearChange >= 0 Then
                  ws.Cells(summaryIterator, 10).Interior.ColorIndex = 4 'green
             Else
                  ws.Cells(summaryIterator, 10).Interior.ColorIndex = 3 'red
             End If
             If firstOpenPrice <> 0 Then
                  pctChange = yearChange / firstOpenPrice
             Else 'it's zero, avoid div/0 error
                  pctChange = 0
             End If
             ws.Cells(summaryIterator, 11).Value = pctChange
             ws.Cells(summaryIterator, 12).Value = totalVol
             
             'update the "mosts" variables if current ticker larger than stored values
             If totalVol > mostVol Then
                 mostVol = totalVol
                 mostVolT = currentRow
             End If
             ' check for biggest + incr
             If pctChange > 0 And pctChange > biggestIncr Then
                   biggestIncrT = currentRow
                   biggestIncr = pctChange
             End If
             ' check for biggest decr
             If pctChange < 0 And pctChange < biggestDecr Then
                   biggestDecrT = currentRow
                   biggestDecr = pctChange
             End If
             
             summaryIterator = summaryIterator + 1
             'update curr ticket's values will start over on next iteration w/next ticker's vals
             totalVol = 0
             curTicker = ""
   
           End If
       Next i 'end row loop
       
       'Label and Format sheet columns
       ws.Cells(1, 9).Value = "Ticker"
       ws.Cells(1, 10).Value = "Yearly Change"
       ws.Cells(1, 11).Value = "Percent Change"
       ws.Cells(1, 12).Value = "Total Stock Volume"
       ws.Cells(1, 15).Value = "Ticker"
       ws.Cells(1, 16).Value = "Value"
       ws.Cells(2, 14).Value = "Biggest Increase (%)"
       ws.Cells(3, 14).Value = "Biggest Decrease (%)"
       ws.Cells(4, 14).Value = "Highest Total Vol."
       ws.Range("K:K").NumberFormat = "0.00%"
       ws.Columns("L").AutoFit
       ws.Columns("N").AutoFit
       
       'Update the "mosts" box
       ws.Cells(2, 15).Value = biggestIncrT
       ws.Cells(2, 16).Value = biggestIncr
       ws.Cells(2, 16).NumberFormat = "0.00%"
       ws.Cells(3, 15).Value = biggestDecrT
       ws.Cells(3, 16).Value = biggestDecr
       ws.Cells(3, 16).NumberFormat = "0.00%"
       ws.Cells(4, 15).Value = mostVolT
       ws.Cells(4, 16).Value = mostVol
   Next ws  'end sheet loop

End Sub


