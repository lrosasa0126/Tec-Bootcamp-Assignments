Sub WallStreet()

Dim Column As Integer
Dim Ticker As String
Dim Counter As Integer
Dim Row As Long
Dim TotalStock As Double
Dim OpenPeriod As Double
Dim ClosePeriod As Double

Row = Cells(Rows.Count, "a").End(xlUp).Row
Counter = 0
TotalStock = 0

'Get the the ticker symbol.

   For i = 2 To Row
  
      If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            Ticker = Cells(i, 1).Value
            Counter = Counter + 1
            Range("I" & Counter + 1).Value = Ticker
            
            
            
 'Get the yearly change

             OpenPeriod = Cells(i, 3).Value
              ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
              ClosePeriod = Cells(i, 6).Value
              Range("J" & Counter + 1).Value = ClosePeriod - OpenPeriod
       
  'Get the Percent Change
       
              YearlyChange = Range("J" & Counter + 1).Value
              Range("k:K").NumberFormat = "0.00%"
              Range("K" & Counter + 1).Value = (YearlyChange / OpenPeriod)
       
   'Get the Total Stock
             
              
              TotalStock = TotalStock + (Cells(i, 7).Value)
              Range("L" & Counter + 1).Value = TotalStock
              TotalStock = 0
              
              Else
              
              TotalStock = TotalStock + (Cells(i, 7).Value)
              
            
       End If
       
       
  Next i
  
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
             
End Sub

