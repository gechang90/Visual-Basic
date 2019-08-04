Attribute VB_Name = "ticker"
Sub ticker()

Dim LastRow As Long
Dim i As Long
Dim count As Long
Dim open_price As Double
Dim close_price As Double
Dim percent_change As Double
Dim sum As Double

sum = 0

LastRow = Cells(Rows.count, 1).End(xlUp).Row

open_price = Cells(2, 3).Value

For i = 2 To LastRow

   If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
   count = count + 1
   Cells(count + 1, 9).Value = Cells(i, 1).Value
   close_price = Cells(i, 6).Value
   Cells(count + 1, 10).Value = close_price - open_price
   percent_change = (Cells(count + 1, 10).Value) / (open_price) * 100
   Cells(count + 1, 11).Value = percent_change
   open_price = Cells(i + 1, 3).Value
   Cells(count + 1, 12).Value = sum + Cells(i, 7).Value
   sum = 0
   
   Else
   sum = sum + Cells(i, 7).Value
   
  End If
 Next i
      
End Sub
