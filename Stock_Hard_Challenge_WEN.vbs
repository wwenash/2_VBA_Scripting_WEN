Sub StocksHardChallenge():

'WEN - VBA Homework - WASHSTL201809DATA3
'set dimensions and declare variables
Dim rowcount As Long
Dim rowStart As Double
Dim total As Double
Dim i As Long
Dim rowTracker As Integer
Dim openPrice As Double
Dim closePrice As Double
Dim priceChange As Double
Dim percentChange As Double
Dim tickerHeader As String
Dim yChangeHeader As String
Dim pChangHeader As String
Dim totVolHeader As String
Dim rowLast As Double
Dim maxValue As Double
Dim minValue As Double
Dim maxVolume As Double
Dim maxRowNumber As Double
Dim minRowNumber As Double
Dim maxVolRowNumber As Double
Dim greatIncrease As String
Dim greatDecrease As String
Dim greatTotVol As String
Dim valueHeader As String
Dim ws As Worksheet

For Each ws In Worksheets

'Set variables for each worksheet

tickerHeader = "Ticker"
yChangeHeader = "Yearly Change"
pChangHeader = "Percent Change"
totVolHeader = "Total Stock Volume"
greatIncrease = "Greatest % Increase"
greatDecrease = "Greatest % Decrease"
greatTotVol = "Greatest Total Volume"
valueHeader = "Value"
i = 0
rowTracker = 0
rowStart = 2
rowcount = 0
total = 0
rowLast = 0

'Add header to column
ws.Range("I1").Value = tickerHeader
ws.Range("J1").Value = yChangeHeader
ws.Range("K1").Value = pChangHeader
ws.Range("L1").Value = totVolHeader
ws.Range("O2").Value = greatIncrease
ws.Range("O3").Value = greatDecrease
ws.Range("O4").Value = greatTotVol
ws.Range("P1").Value = tickerHeader
ws.Range("Q1").Value = valueHeader

' get the row number of the last row with data
rowcount = ws.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowcount
    'Detect change in ticker, write out total and stock name
    If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
        
        ' Find First non zero starting value
            If ws.Cells(rowStart, 3) = 0 Then
                Dim findPosValue As Variant
            
                For findPosValue = rowStart To i
                    If ws.Cells(findPosValue, 3).Value <> 0 Then
                     
                        rowStart = findPosValue
                        Exit For
                    End If
                 Next findPosValue
            End If
            
        'Calculate change in stock price
        openPrice = ws.Cells(rowStart, 3).Value
        closePrice = ws.Cells(i, 6).Value
        
        priceChange = closePrice - openPrice
        percentChange = Round((priceChange / openPrice) * 100, 2)
        
        ' Store and print results
        total = total + ws.Cells(i, 7).Value
        ws.Range("I" & 2 + rowTracker).Value = ws.Cells(i, 1).Value
        ws.Range("J" & 2 + rowTracker).Value = priceChange
        ws.Range("K" & 2 + rowTracker).Value = percentChange & "%"
        ws.Range("L" & 2 + rowTracker).Value = total
        
        ' color formating... if positive make green and if negative make red
        If (priceChange >= 0) Then
            ws.Range("J" & 2 + rowTracker).Interior.ColorIndex = 4
        Else
            ws.Range("J" & 2 + rowTracker).Interior.ColorIndex = 3
        End If
        
        'This tracker keeps count where the results get printed
        rowTracker = rowTracker + 1
        'Reset values
        total = 0
        'Start of next stock
        rowStart = i + 1
        
    Else
        'keep running total of Total Stock Volume
        total = total + ws.Cells(i, 7).Value
        
    End If
Next i

 'get the row number of the last row with data
rowLast = ws.Cells(Rows.Count, "I").End(xlUp).Row

'Find max, min percentage and greatest volume
maxValue = WorksheetFunction.Max(ws.Range("K2:K" & rowLast)) * 100
minValue = WorksheetFunction.Min(ws.Range("K2:K" & rowLast)) * 100
maxVolume = WorksheetFunction.Max(ws.Range("L2:L" & rowLast))

'Match max, min percentage and greatest volume with the row number
maxRowNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowLast)), ws.Range("K2:K" & rowLast), 0)
minRowNumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowLast)), ws.Range("K2:K" & rowLast), 0)
maxVolRowNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowLast)), ws.Range("L2:L" & rowLast), 0)

'Print out results
ws.Range("Q2").Value = maxValue & "%"
ws.Range("Q3").Value = minValue & "%"
ws.Range("Q4").Value = maxVolume
ws.Range("P2") = ws.Cells(maxRowNumber + 1, 9)
ws.Range("P3").Value = ws.Cells(minRowNumber + 1, 9)
ws.Range("P4").Value = ws.Cells(maxVolRowNumber + 1, 9)

'Autofit column width
ws.Range("J1:Q1").EntireColumn.AutoFit

Next ws

End Sub