Attribute VB_Name = "Module1"
Sub testing():

Dim Ticker As String
Dim Total_Stock_Volume As Double
Dim Row_Total As Long
Dim LastRow As Long
Dim i As Long

'Print column headings
Range("J1").Value = "Ticker"
Range("K1").Value = "Total Stock Volume"

'Set initial value of Row_Total
Row_Total = 2

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
'Where to find Ticker
Ticker = Cells(i, 1).Value
'Where to find volume and add them
Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
'Where to print Ticker
Range("J" & Row_Total).Value = Ticker
'Where to print volume
Range("K" & Row_Total).Value = Total_Stock_Volume

'Increment Row_Total to next row until LastRow is reached
Row_Total = Row_Total + 1

'Initial value of volume
Total_Stock_Volume = 0

Else
Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

End If

Next i

End Sub







