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

Dim stockPointer As Long
Dim i As Long
Dim stockTicker  As String
Dim stockOpen As Long
Dim stockClose As Long
Dim stockVolume As Long
Dim lastRow As Long


lastRow = Cells(Rows.Count, "A").End(xlUp).Row + 1


For i = 2 To lastRow

    If Cells(i, 1).Value <> stockTicker Then
    stockClose = Cells(i - 1, 6)
    
    stockTicker = Cells(i, 1)
    
    End If
    

Next i

MsgBox (stockClose)

End Sub
