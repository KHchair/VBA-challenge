VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_scraper()

Dim stock_array_increment As Long
Dim last_row As Long
Dim i As Long
Dim stock_ticker(0 To 9999) As String
Dim stock_open(0 To 9999) As Long
Dim stock_close(0 To 9999) As Long
Dim stock_percent_change(0 To 9999) As Long
Dim stock_volume(0 To 9999) As Long
Dim closingprice As Long
Dim openingprice As Long



stock_array_increment = 0
last_row = Cells(Rows.Count, "A").End(xlUp).Row

'gets a mismatch error out of the way
Cells(1, 6).Value = 0

'for loop to scrape data needed for data dump on sheet
For i = 2 To last_row
    
    'ticker logic
     If Cells(i, 1).Value <> stock_ticker(stock_array_increment) Then
    ' stock_close(stock_array_increment) = Cells(i - 1, 6).Value
    ' closingprice = stock_close(stock_array_increment)
     
     'if statement to prevent overflow error from first row of data
     '   If closingprice = 0 Then
     '   closingprice = 999
     '   End If
     
     'openingprice = stock_open(stock_array_increment)
     'stock_percent_change(stock_array_increment) = ((closingprice - openingprice) / closingprice)
     'MsgBox (stock_array_increment)
     stock_array_increment = stock_array_increment + 1
     'stock_ticker(stock_array_increment) = Cells(i, 1).Value
     'stock_open(stock_array_increment) = Cells(i, 3).Value
     'stock_volume(stock_array_increment) = (stock_volume(stock_array_increment) + Cells(i, 7).Value)
     
     Else
     
     stock_volume(stock_array_increment) = stock_volume(stock_array_increment) + Cells(i, 7).Value
     
     End If
     
Next i

'Print scraped data into excel sheet
For i = 1 To stock_array_increment
    Cells(i, 9) = stock_ticker(i)
    Cells(i, 10) = stock_close(i) - stock_open(i)
    Cells(i, 11) = stock_percent_change(i)
    Cells(i, 12) = stock_volume(i)

Next i

End Sub
