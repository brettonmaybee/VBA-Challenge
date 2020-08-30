Attribute VB_Name = "Module1"
Sub Stock_Data()
  Cells(1, "i") = "Ticker"
  Cells(1, "j") = "Yearly Change"
  Cells(1, "k") = "% Change"
  Cells(1, "l") = "Total Stock Volume"
  Cells(1, "o") = "in_price"
  Cells(1, "p") = "fin_price"
  
  ' Set an initial variable for holding the brand name
  Dim Ticker As String

  ' Set an initial variable for holding the total value of each stock
  Dim Stock_Total As Double
  Stock_Total = 0
  
  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  Dim conter As Integer
  counter = 0
  
  Dim in_price As Double
  
  Dim fin_price As Double
   
  Dim ann_change As Variant
  
  ' Loop through all The Data
    For i = 2 To 705719
  
  'Check if we are still within the same stock, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

  ' Set the Ticker name
    Ticker = Cells(i, 1).Value

  ' Add to the stock total
    Stock_Total = Stock_Total + Cells(i, 7).Value
    
  'Add the Final Price
      fin_price = Cells(i, 6).Value
      
      in_price = Cells(i - counter, 6).Value
      
      ann_change = fin_price - in_price
      
      If in_price <> 0 Then
        
        per_change = (ann_change / in_price) * 0.01
        
        
      End If
      
      ' Print the Ticker in the Summary Table
      Range("i" & Summary_Table_Row).Value = Ticker

      ' Print the stock total to the Summary Table
      Range("l" & Summary_Table_Row).Value = Stock_Total

      ' Print the final price in the Summary Table
      Range("P" & Summary_Table_Row).Value = fin_price
      
      Range("O" & Summary_Table_Row).Value = in_price
      
      Range("K" & Summary_Table_Row).Value = per_change
      
      Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
      Range("J" & Summary_Table_Row).Value = ann_change
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      Stock_Total = 0
        
      counter = 0
      ' If the cell immediately following a row is the same ticker
    Else

      ' Add to the Stock Value to the stock total
      Stock_Total = Stock_Total + Cells(i, 7).Value
          
      counter = counter + 1
    End If

  Next i

End Sub
