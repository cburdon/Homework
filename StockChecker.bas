Attribute VB_Name = "Module1"
Sub StockActual()
'Declare variables to track ticker, total volume, last row, result row,etc.
    Dim ticker As String
    Dim totalVol As Double
    Dim lastRow As Long
    Dim tickerCounter As Long
    Dim yearOpen As Double
    Dim yearClose As Double
    Dim change As Double
    Dim percentChange
    Dim maxPercentChangePlus
    Dim maxPercentChangeMinus
    Dim maxTicker As String
    Dim minTicker As String
    Dim maxVol As Double
    Dim maxVolTicker As String
    
    maxPercentChangePlus = 0
    maxPercentChangeMinus = 0
    maxVol = 0
    tickerCounter = 2
    
    
    
'iterate through all stock worksheets
    For Each ws In Worksheets
    'set up column and row labels for new chart
    
        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Yearly Change"
        ws.Cells(1, 10).Value = "Percent Change"
        ws.Cells(1, 11).Value = "Total Stock Vol"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % change positive"
        ws.Cells(3, 15).Value = "Greatest % change negative"
        ws.Cells(4, 15).Value = "Greatest total Volume"
    'determine last row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'iterate through one worksheet
        
        For i = 2 To lastRow
        'check to see if new ticker, pull ticker, totalVol, change, percentChange,
        'and yearOpen. Record on master table. Format Change column by number sign.
                If ws.Cells(i, 1).Value <> ticker Then
                    ticker = ws.Cells(i, 1).Value
                    ws.Cells(tickerCounter, 8).Value = ticker
                    ws.Cells(tickerCounter, 11).Value = totalVol
                    totalVol = ws.Cells(i, 7).Value
                    yearOpen = ws.Cells(i, 3).Value
                    ws.Cells(tickerCounter, 9).Value = change
                    ws.Cells(tickerCounter, 9).NumberFormat = "0.000000000"
                    ws.Cells(tickerCounter, 10).Value = percentChange
                    ws.Cells(tickerCounter, 10).NumberFormat = "0.000%"
                    'increase ticker reference variable
                    tickerCounter = tickerCounter + 1
                'if not new ticker, pull totalVol and year close and get change and
                'percent change calculated
                Else
                    totalVol = ws.Cells(i, 7).Value + totalVol
                    ws.Cells(tickerCounter - 1, 11).Value = totalVol
                    yearClose = ws.Cells(i, 6).Value
                    change = yearClose - yearOpen
                    If change > 0 Then
                        ws.Cells(tickerCounter - 1, 9).Interior.ColorIndex = 4
                    Else
                        ws.Cells(tickerCounter - 1, 9).Interior.ColorIndex = 3
                    End If
                    ws.Cells(tickerCounter - 1, 9).Value = change
                    'prevent undefined percentChange calculation
                    If yearOpen <> 0 Then
                        percentChange = change / yearOpen
                        ws.Cells(tickerCounter - 1, 10).Value = percentChange
                    Else
                        ws.Cells(tickerCounter - 1, 10).Value = "N/A"
                    End If
                    
                    
                End If
        Next i
        'iterate through final list of values to pull max pos and neg percent change and greatest total vol
        
     For j = 2 To tickerCounter
        
        If ws.Cells(j, 10).Value <> "N/A" Then
            percentChange = ws.Cells(j, 10).Value
        Else
            percentChange = 0
        End If
        
        totalVol = ws.Cells(j, 11).Value
    
        If percentChange > maxPercentChangePlus Then
            maxPercentChangePlus = percentChange
            maxTicker = ws.Cells(j, 8).Value
        ElseIf percentChange < maxPercentChangeMinus Then
            maxPercentChangeMinus = percentChange
            minTicker = ws.Cells(j, 8).Value
        End If
        
        If totalVol > maxVol Then
            maxVol = totalVol
            maxVolTicker = ws.Cells(j, 8).Value
        Else
        End If
     Next j
    
    'create separate chart for those values and format it
    
    ws.Cells(2, 16).Value = maxTicker
    ws.Cells(2, 16).Interior.ColorIndex = 4
    ws.Cells(3, 16).Value = minTicker
    ws.Cells(3, 16).Interior.ColorIndex = 3
    ws.Cells(2, 17).Value = maxPercentChangePlus
    ws.Cells(2, 17).NumberFormat = "0.000%"
    ws.Cells(2, 17).Interior.ColorIndex = 4
    ws.Cells(3, 17).Value = maxPercentChangeMinus
    ws.Cells(3, 17).NumberFormat = "0.000%"
    ws.Cells(3, 17).Interior.ColorIndex = 3
    ws.Cells(4, 16).Value = maxVolTicker
    ws.Cells(4, 17).Value = maxVol
    
    'reset values for next worksheet
    
    maxPercentChangePlus = 0
    maxPercentChangeMinus = 0
    maxVol = 0
    tickerCounter = 2
    
    
    
    
    Next ws
    
        
                        
                    
                        
            
    
    
            
                    
                    
                
                
                
        





End Sub





