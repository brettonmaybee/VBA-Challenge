Attribute VB_Name = "Module11"
Sub Stock_Data()

'Dim WS_Count As Integer

'Dim X As Integer

'WS_Count = ActiveWorkbook.Worksheets.Count

'For X = 1 To WS_Count

'Worksheets(X).Activate
            
  Cells(1, "i") = "Ticker"
  Cells(1, "j") = "Yearly Change"
  Cells(1, "k") = "% Change"
  Cells(1, "l") = "Total Stock Volume"
  Cells(1, "o") = "in_price"
  Cells(1, "p") = "fin_price"
  
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
  
  
    For i = 2 To 800000
  
  
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

   
    Ticker = Cells(i, 1).Value

  
    Stock_Total = Stock_Total + Cells(i, 7).Value
    
  
      fin_price = Cells(i, 6).Value
      
      in_price = Cells(i - counter, 6).Value
      
      ann_change = fin_price - in_price
      
    If in_price <> 0 Then
        
        per_change = (ann_change / in_price) * 0.01
        
        
    End If
       
      Range("i" & Summary_Table_Row).Value = Ticker

      Range("l" & Summary_Table_Row).Value = Stock_Total

      Range("P" & Summary_Table_Row).Value = fin_price
      
      Range("O" & Summary_Table_Row).Value = in_price
      
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

For j = 2 To 800000
  
  If Cells(j, 11) > 0 Then
        Cells(j, 11).Interior.ColorIndex = 4
    Else
        Cells(j, 11).Interior.ColorIndex = 3
  End If
  
  Next j
    
    'Next X

  max = Application.WorksheetFunction.max(Range("j2: j800000"))
     
     Cells(1, 17).Value = max
  
  Min = Application.WorksheetFunction.Min(Range("j2:j800000"))
      
     Cells(1, 18).Value = Min
      
    End Sub
