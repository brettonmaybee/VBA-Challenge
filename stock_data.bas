Attribute VB_Name = "Module11"
Sub Stock_Data()

Dim WS_Count As Integer

Dim X As Integer

WS_Count = ActiveWorkbook.Worksheets.Count

For X = 1 To WS_Count

Worksheets(X).Activate
            
  Cells(1, "i") = "Ticker"
  Cells(1, "j") = "Yearly Change"
  Cells(1, "k") = "% Change"
  Cells(1, "l") = "Total Stock Volume"
  Cells(1, "p") = "Ticker"
  Cells(1, "q") = "Value"
  Cells(2, "o") = "Greatest % Increase"
  Cells(3, "o") = "Greatest % Decrease"
  Cells(4, "o") = "Greatest Total Volume"
  
Dim end_data As Long

Dim end_summ As Integer

end_data = Cells(Rows.Count, 1).End(xlUp).Row


Dim Ticker As String

Dim Stock_Total As Double
  
  Stock_Total = 0
  
Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
Dim conter As Integer
  counter = 0
  
Dim in_price As Double
  
Dim fin_price As Double
   
Dim ann_change As Variant
  
  
    For i = 2 To end_data
  
  
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

   
    Ticker = Cells(i, 1).Value

  
    Stock_Total = Stock_Total + Cells(i, 7).Value
    
  
      fin_price = Cells(i, 6).Value
      
      in_price = Cells(i - counter, 6).Value
      
      ann_change = fin_price - in_price
      
    If in_price <> 0 Then
        
        per_change = (ann_change / in_price)
        
        
    End If
       
      Range("i" & Summary_Table_Row).Value = Ticker

      Range("l" & Summary_Table_Row).Value = Stock_Total
      
      Range("K" & Summary_Table_Row).Value = per_change
      
      Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
      Range("J" & Summary_Table_Row).Value = ann_change
      
    
      Summary_Table_Row = Summary_Table_Row + 1
      
      Stock_Total = 0
        
      counter = 0
      
    Else

      Stock_Total = Stock_Total + Cells(i, 7).Value
          
      counter = counter + 1
    End If

  Next i

end_summ = Cells(Rows.Count, "i").End(xlUp).Row

For j = 2 To end_summ
  
  If Cells(j, 11) > 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
        
    Else
        Cells(j, 10).Interior.ColorIndex = 3
  End If
  
  Next j
    
  Summer_Table_Row = 0
  
  
  Max = Application.WorksheetFunction.Max(Range("k2:k3000"))
     Cells(2, "q").NumberFormat = "0.00%"
     Cells(2, "q").Value = Max
  
  Min = Application.WorksheetFunction.Min(Range("k2:k3000"))
     Cells(3, "q").NumberFormat = "0.00%"
     Cells(3, "q").Value = Min
      
  max_vol = Application.WorksheetFunction.Max(Range("l2:k3000"))
    Cells(4, "q") = max_vol


For k = 2 To 3000
  
  If Cells(k, "k").Value = Cells(2, "q").Value Then

    Cells(2, "p").Value = Cells(k, "i").Value
  
  End If
  
  If Cells(k, "k").Value = Cells(3, "q").Value Then
  
    Cells(3, "p").Value = Cells(k, "i").Value
  
  End If
  
  If Cells(k, "l").Value = Cells(4, "q").Value Then
  
    Cells(4, "p").Value = Cells(k, "i").Value
  
  End If
  
Next k
    
  Next X

  End Sub
