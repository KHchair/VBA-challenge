VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_scraper_v2()

Dim i As Long
Dim j As Long
Dim stockTicker  As String
Dim stockOpen As Long
Dim stockClose As Long
Dim stockVolume As Double
Dim lastRow As Long

'set "<close>" header to 0 to avoid a data mismatch on first loop
Cells(1, 6) = 0

'set lastRow var to +1 to accomodate algo logic for the last stock
lastRow = Cells(Rows.Count, "A").End(xlUp).Row + 1
j = 1

'passes through every row in the data set
For i = 2 To lastRow

    'looks for stock name transitions and executes necessary operations
    If Cells(i, 1).Value <> stockTicker Then
    stockClose = Cells(i - 1, 6)
    
    'print data stored in vars to the sheet
    Cells(j, 9) = stockTicker
    Cells(j, 10) = stockClose - stockOpen
        
        'fills green for increased price, red for decreased price, none for 0
        If stockClose > stockOpen Then
        Cells(j, 10).Interior.ColorIndex = 4
        ElseIf stockClose < stockOpen Then
        Cells(j, 10).Interior.ColorIndex = 3
        End If
        
        'work around for reading 0 from the first row on initial loop
        'prints percent change
        If stockOpen <> 0 Then
        Cells(j, 11) = (stockClose - stockOpen) / stockOpen
        End If
    
    Cells(j, 12) = stockVolume
      
    'changes variables for next data print line
    stockTicker = Cells(i, 1)
    stockOpen = Cells(i, 3)
    stockVolume = 0
    j = j + 1
    
   
    'tallies volumes if  stockTicker is equal to prior row's stock
    Else
    
    stockVolume = stockVolume + Cells(i, 7)
        
    End If

Next i

'restores header to pre sub running
Cells(1, 6) = "<close>"

'Adds headers for newly populated data
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Price Chang"
Cells(1, 11) = "Percent Price Change"
Cells(1, 12) = "Total Volume"

'challenge
Dim greatestInc As Long
Dim greatestDec As Long
Dim greatestVol As Double

'print table headers
Cells(2, 15) = "Greatest % Increase"
Cells(3, 15) = "Greatest % Decrease"
Cells(4, 15) = "Greatest Total Volume"
Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"

'set lastRow based on first column of table populated by this subroutine
lastRow = Cells(Rows.Count, "I").End(xlUp).Row
stockTicker = Cells(2, 9)
greatestInc = Cells(2, 11)
greatestDec = Cells(2, 11)
greatestVol = Cells(2, 12)

'set cell values incase the first stock is the rightful holder of any category
Cells(2, 16) = stockTicker
Cells(2, 17) = greatestInc
Cells(3, 16) = stockTicker
Cells(3, 17) = greatestDec
Cells(4, 16) = stockTicker
Cells(4, 17) = greatestVol

'compares new 3 variables on each increment of the table to see who is "most" of their respective category
For i = 2 To lastRow

    If greatestInc < Cells(i, 11) Then
    greatestInc = Cells(i, 11)
    stockTicker = Cells(i, 9)
    Cells(2, 16) = stockTicker
    Cells(2, 17) = greatestInc
    End If
    
    If greatestDec > Cells(i, 11) Then
    greatestDec = Cells(i, 11)
    stockTicker = Cells(i, 9)
    Cells(3, 16) = stockTicker
    Cells(3, 17) = greatestDec
    End If
    
    If greatestVol < Cells(i, 12) Then
    greatestVol = Cells(i, 12)
    stockTicker = Cells(i, 9)
    Cells(4, 16) = stockTicker
    Cells(4, 17) = greatestVol
    End If
    
Next i

End Sub

