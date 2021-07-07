VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_scraper_v2()

Dim i As Double
Dim j As Double
Dim stockTicker  As String
Dim stockOpen As Double
Dim stockClose As Double
Dim stockVolume As Double
Dim lastRow As Double

'Cycles through every worksheet in excel file
For Each ws In Worksheets

'set "<close>" header to 0 to avoid a data mismatch on first loop
ws.Cells(1, 6) = 0

'set lastRow var to +1 to accomodate algo logic for the last stock
lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
j = 1

'passes through every row in the data set
For i = 2 To lastRow

    'looks for stock name transitions and executes necessary operations
    If ws.Cells(i, 1).Value <> stockTicker Then
    stockClose = ws.Cells(i - 1, 6)
    
    'print data stored in vars to the sheet
    ws.Cells(j, 9) = stockTicker
    ws.Cells(j, 10) = stockClose - stockOpen
        
        'fills green for increased price, red for decreased price, none for 0
        If stockClose > stockOpen Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
        ElseIf stockClose < stockOpen Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
        End If
        
        'work around for reading 0 from the first row on initial loop
        'prints percent change
        If stockOpen <> 0 Then
        ws.Cells(j, 11) = (stockClose - stockOpen) / stockOpen
        End If
    
    ws.Cells(j, 12) = stockVolume
      
    'changes variables for next data print line
    stockTicker = ws.Cells(i, 1)
    stockOpen = ws.Cells(i, 3)
    stockVolume = 0
    j = j + 1
    
   
    'tallies volumes if  stockTicker is equal to prior row's stock
    Else
    
    stockVolume = stockVolume + ws.Cells(i, 7)
        
    End If

Next i

'restores header to pre sub running
ws.Cells(1, 6) = "<close>"

'Adds headers for newly populated data
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Price Chang"
ws.Cells(1, 11) = "Percent Price Change"
ws.Cells(1, 12) = "Total Volume"

'challenge
Dim greatestInc As Long
Dim greatestDec As Long
Dim greatestVol As Double

'print table headers
ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(3, 15) = "Greatest % Decrease"
ws.Cells(4, 15) = "Greatest Total Volume"
ws.Cells(1, 16) = "Ticker"
ws.Cells(1, 17) = "Value"

'set lastRow based on first column of table populated by this subroutine
lastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
stockTicker = ws.Cells(2, 9)
greatestInc = ws.Cells(2, 11)
greatestDec = ws.Cells(2, 11)
greatestVol = ws.Cells(2, 12)

'set cell values incase the first stock is the rightful holder of any category
ws.Cells(2, 16) = stockTicker
ws.Cells(2, 17) = greatestInc
ws.Cells(3, 16) = stockTicker
ws.Cells(3, 17) = greatestDec
ws.Cells(4, 16) = stockTicker
ws.Cells(4, 17) = greatestVol

'compares new 3 variables on each increment of the table to see who is "most" of their respective category
For i = 2 To lastRow

    If greatestInc < ws.Cells(i, 11) Then
    greatestInc = ws.Cells(i, 11)
    stockTicker = ws.Cells(i, 9)
    ws.Cells(2, 16) = stockTicker
    ws.Cells(2, 17) = greatestInc
    End If
    
    If greatestDec > ws.Cells(i, 11) Then
    greatestDec = ws.Cells(i, 11)
    stockTicker = ws.Cells(i, 9)
    ws.Cells(3, 16) = stockTicker
    ws.Cells(3, 17) = greatestDec
    End If
    
    If greatestVol < ws.Cells(i, 12) Then
    greatestVol = ws.Cells(i, 12)
    stockTicker = ws.Cells(i, 9)
    ws.Cells(4, 16) = stockTicker
    ws.Cells(4, 17) = greatestVol
    End If
    
Next i

Next ws

End Sub
