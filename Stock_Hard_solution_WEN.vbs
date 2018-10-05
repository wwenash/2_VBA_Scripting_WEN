Sub StocksHard():

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
Range("I1").value = tickerHeader
Range("J1").value = yChangeHeader
Range("K1").value = pChangHeader
Range("L1").value = totVolHeader
Range("O2").value = greatIncrease
Range("O3").value = greatDecrease
Range("O4").value = greatTotVol
Range("P1").value = tickerHeader
Range("Q1").value = valueHeader

' get the row number of the last row with data
rowcount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowcount
    'Detect change in ticker, write out total and stock name
    If (Cells(i + 1, 1).value <> Cells(i, 1).value) Then
    
         ' Find First non zero starting value
        If Cells(rowStart, 3) = 0 Then
            Dim findPosValue As Variant            
            For findPosValue = rowStart To i
                If Cells(findPosValue, 3).Value <> 0 Then
                     
                    rowStart = findPosValue
                    Exit For
                End If
                Next findPosValue
        End If

        'Calculate change in stock price
        openPrice = Cells(rowStart, 3).value
        closePrice = Cells(i, 6).value
        
        priceChange = closePrice - openPrice
        percentChange = Round((priceChange / openPrice) * 100, 2)
        
        ' Store and print results
        total = total + Cells(i, 7).value
        Range("I" & 2 + rowTracker).value = Cells(i, 1).value
        Range("J" & 2 + rowTracker).value = priceChange
        Range("K" & 2 + rowTracker).value = percentChange & "%"
        Range("L" & 2 + rowTracker).value = total
        
        ' color formating... if positive make green and if negative make red
        If (priceChange >= 0) Then
            Range("J" & 2 + rowTracker).Interior.ColorIndex = 4
        Else
            Range("J" & 2 + rowTracker).Interior.ColorIndex = 3
        End If
        
        'This tracker keeps count where the results get printed
        rowTracker = rowTracker + 1
        'Reset values
        total = 0
        'Start of next stock
        rowStart = i + 1
        
    Else
        'keep running total of Total Stock Volume
        total = total + Cells(i, 7).value
        
    End If
Next i

 'get the row number of the last row with data
rowLast = Cells(Rows.Count, "I").End(xlUp).Row

'Find max, min percentage and greatest volume
maxValue = WorksheetFunction.Max(Range("K2:K" & rowLast)) * 100
minValue = WorksheetFunction.Min(Range("K2:K" & rowLast)) * 100
maxVolume = WorksheetFunction.Max(Range("L2:L" & rowLast))

'Match max, min percentage and greatest volume with the row number
maxRowNumber = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowLast)), Range("K2:K" & rowLast), 0)
minRowNumber = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowLast)), Range("K2:K" & rowLast), 0)
maxVolRowNumber = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowLast)), Range("L2:L" & rowLast), 0)

'Print out results
Range("Q2").value = maxValue & "%"
Range("Q3").value = minValue & "%"
Range("Q4").value = maxVolume
Range("P2") = Cells(maxRowNumber + 1, 9)
Range("P3").value = Cells(minRowNumber + 1, 9)
Range("P4").value = Cells(maxVolRowNumber + 1, 9)

'Autofit column width
Range("J1:Q1").EntireColumn.AutoFit

End Sub
