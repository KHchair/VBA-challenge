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

Dim stockPointer As Long
Dim i As Long
Dim stockTicker  As String
Dim stockOpen As Long
Dim stockClose As Long
Dim stockVolume As Long
Dim lastRow As Long
Dim j As Long


stockOpen = Cells(2, 3)
Cells(1, 6) = 0
stockClose = 0
lastRow = Cells(Rows.Count, "A").End(xlUp).Row + 1
j = 1

For i = 2 To lastRow

    If Cells(i, 1).Value <> stockTicker Then
    stockClose = Cells(i - 1, 6)
    
    
    Cells(j, 9) = stockTicker
    Cells(j, 10) = stockClose - stockOpen
        
        If stockClose <> 0 Then
        Cells(j, 11) = (stockClose - stockOpen) / stockClose
        End If
        
    'Cells(j, 11) = (stockClose - stockOpen) / stockClose
    
    'changes variables for next data print line
    stockTicker = Cells(i, 1)
    stockOpen = Cells(i, 3)
    stockVolume = 0
    j = j + 1
    
   
    
    Else
    
    
    'stockVolume = stockVolume + Cells(i, 7)
    
    
    
    End If
    

Next i

'restores header to pre sub running
Cells(1, 6) = "<close>"

'Adds headers for newly populated data
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Price Chang"
Cells(1, 11) = "Percent Price Change"
Cells(1, 12) = "Total Volume"

MsgBox (stockClose)

End Sub

