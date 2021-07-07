VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_scraper()

Dim stock_array_increment As Long
Dim last_row As Long
Dim i As Long
Dim stock_ticker() As String
Dim stock_open() As Long
Dim stock_close() As Long
Dim stock_percent_change() As Long
Dim stock_volume() As Long


stock_array_increment = 0
last_row = Cells(Rows.Count, "A").End(xlUp).Row

For i = 1 To last_row
    
    'ticker logic
     If Cells(i, 1).Value <> stock_ticker(stock_array_increment - 1) Then
     stock_close(stock_array_increment) = Cells(i - 1, 6).Value
     stock_percent_change(stock_array_increment) = (stock_close - stock_open) / stock_open
     stock_array_increment = stock_array_increment + 1
     stock_ticker(stock_array_increment) = Cells(i, 1).Value
     stock_open(stock_array_increment) = Cells(i, 3).Value
     
     
     
     Elseif
     
     End If
     

    

    



End Sub

