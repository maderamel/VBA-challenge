Public Sub stockDataChallenge()
 Dim ws As Worksheet
    
    'to loop through all worksheets
    For Each ws In Worksheets
        
        'add headers for ticker, yearly change, percent change, total stock volume
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'define variables
        Dim ticker As String
            ticker = " "
        Dim stockVolume As Double
            stockVolume = 0
        Dim openPrice As Double
            openPrice = 0
        Dim closePrice As Double
            closePrice = 0
        Dim changePercent As Double
            changePercent = 0
        Dim yearlyChange As Double
            yearlyChange = 0
        
            
        'define rows and columns variable and last row of worksheet, i is longbecause large # of rows; j is integer because there's only a small # columns
        Dim i As Long
        Dim j As Integer
        Dim lastrow As Long
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'added ticker row to keep track of rows for ticker symbols
        Dim tickerRow As Long
            tickerRow = 1
        
        'create loop to last row
        For i = 2 To lastrow
        
            'add stockvolume before if
            stockVolume = stockVolume + ws.Cells(i, 7).Value
                
            'output ticker symbol
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                openPrice = ws.Cells(i, 3).Value
                ticker = ws.Cells(i, 1).Value
                
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              closePrice = ws.Cells(i, 6).Value
              yearlyChange = closePrice - openPrice
              
              
              changePercent = (yearlyChange / openPrice)
                'percent format for change percent
                ws.Columns("K").NumberFormat = "0.00%"
                
                'format column size
                ws.Columns("I:L").EntireColumn.AutoFit
                
                'output values at same time
                tickerRow = tickerRow + 1
                ws.Cells(tickerRow, "I").Value = ticker
                
                'output yearly change, percent change
                ws.Cells(tickerRow, "J").Value = yearlyChange
                ws.Cells(tickerRow, "K").Value = changePercent
                ws.Cells(tickerRow, "L").Value = stockVolume
                
                'reset stock vol
                stockVolume = 0
                
                
            End If
        
        Next i
        
    'format yearly change fills
    Dim yrChng As Range
    Dim g As Long
    Dim cellCount As Long
    Dim color_cell As Range
    
    'Got it to run without ws. but when added ws, run error 1004 fail.
    Set yrChng = ws.Range("J2", ws.Range("J2").End(xlDown))
    cellCount = yrChng.Cells.Count
    
    For g = 1 To cellCount
    Set color_cell = yrChng(g)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
            End With
       End Select
    Next g
        
  'define headers for greater functionality portion of ws
  ws.Cells(2, 15).Value = "Greatest % Increase"
  ws.Cells(3, 15).Value = "Greatest % Decrease"
  ws.Cells(4, 15).Value = "Greatest Total Volume"
  ws.Cells(1, 16).Value = "Ticker"
  ws.Cells(1, 17).Value = "Value"
  
  'define variables for greater functionality
  Dim maxIncrease As Double
  Dim minIncrease As Double
  Dim pctCount As Long
  Dim ttlVolcount As Long
  Dim maxTtlvol As Double
  
  'set last row for ranges
  Dim pctRange As Range
    Set pctRange = ws.Range("K2", ws.Range("K2").End(xlDown))
    pctCount = pctRange.Cells.Count
  Dim ttlVolrange As Range
    Set ttlVolrange = ws.Range("L2", ws.Range("L2").End(xlDown))
    ttlVolcount = ttlVolrange.Cells.Count
  
  'define/find valuesmax&min wrong need to ask bcs ttl vol right
  maxIncrease = WorksheetFunction.Max(pctRange)
  minIncrease = WorksheetFunction.Min(pctRange)
  maxTtlvol = WorksheetFunction.Max(ttlVolrange)
  
  'output values
  ws.Cells(2, 17).Value = maxIncrease
  ws.Cells(3, 17).Value = minIncrease
  ws.Cells(4, 17).Value = maxTtlvol
  ws.Cells(2, 16) = WorksheetFunction.XLookup([Q2], [K:K], [I:I])
  ws.Cells(3, 16).Formula = "=XLOOKUP(Q3,K:K,I:I)"
  ws.Cells(4, 16).Formula = "=XLOOKUP(Q4,L:L,I:I)"
  
  'format column size
  ws.Columns("O:Q").EntireColumn.AutoFit

Next ws
        

End Sub
