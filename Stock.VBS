Sub Ticker_info()

   'Set an initial variable for holding the brand name
  Dim Ticker_Name As String

   'Set an initial variable for holding the total per ticker name
  Dim Total_Stock_volume_Total, openprice, closedprice As Double
  
 'Add to the Stock volume Total
      Total_Stock_volume_Total = 0
      openprice = Cells(2, 3).Value
      
  'Counts the number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row

   'Keep track of the location for each ticker symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

   'Loop through all ticker symbols
  For i = 2 To lastrow

  'Dim ws As Worksheet
  
  'For Each ws In Worksheets
    ' ws.Activate
      
  'Add to the Stock volume Total
      Total_Stock_volume_Total = Total_Stock_volume_Total + Cells(i, 7).Value
      
     'Check if we are still within the same ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

       'Set the Ticker symbol
      Ticker_symbol = Cells(i, 1).Value
      closedprice = Cells(i, "F").Value

       'Print the ticker symbol in the Summary Table
      Range("J" & Summary_Table_Row).Value = Ticker_symbol
      Range("K" & Summary_Table_Row).Value = closedprice - openprice
       
       If Range("K" & Summary_Table_Row).Value > 0 Then
           Range("K" & Summary_Table_Row).Interior.ColorIndex = 4

      Else
         Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
         
     End If
     
     If openprice <> 0 Then
        Range("L" & Summary_Table_Row).Value = FormatPercent((closedprice - openprice) / openprice, 2)
        
     Else
        Range("L" & Summary_Table_Row).Value = Null
    
    End If

      'Print the stock total amount  to the Summary Table
      Range("M" & Summary_Table_Row).Value = Total_Stock_volume_Total

       'Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

       'Reset the Brand Total
      Total_Stock_volume_Total = 0
      
      openprice = Cells(i + 1, 3).Value
    
    End If
    

  Next i

End Sub


