Sub Ticker_info()

   'Set an initial variable for holding the Ticker symbol
  Dim Ticker_Name As String

   'Set an initial variable for holding the total stock volume,the opening and closing price per ticker symbol
  Dim Total_Stock_volume_Total, openprice, closedprice As Double
  
  'Add to the Stock volume Total
      Total_Stock_volume_Total = 0
  'retrieve open price
      openprice = Cells(2, 3).Value
      
  'Counts the number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).row

   'Keep track of the location for each ticker symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
'Set each worksheet as a variable
   Dim ws As Worksheet
'loop through each worksheet
   For Each ws In ActiveWorkbook.Worksheets
   ws.Activate
    
   'Loop through all ticker symbols
   For I = 2 To lastrow

  'Add to the Stock volume Total
    Total_Stock_volume_Total = Total_Stock_volume_Total + Cells(I, 7).Value
      
     'Check if we are still within the same ticker symbol, if it is not...
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

       'Set the Ticker symbol
      Ticker_symbol = Cells(I, 1).Value
       'Set the cloing price
      closedprice = Cells(I, "F").Value

       'Print the ticker symbol in the Summary Table
      Range("J" & Summary_Table_Row).Value = Ticker_symbol
       'Print the yearly change in the summary table
      Range("K" & Summary_Table_Row).Value = closedprice - openprice
       
       'Conditional formating to Colour green or red depending on percentage
      If Range("K" & Summary_Table_Row).Value > 0 Then
           Range("K" & Summary_Table_Row).Interior.ColorIndex = 4

      Else
         Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
         
      End If
     
       'Calculate percentage change
      If openprice <> 0 Then
        Range("L" & Summary_Table_Row).Value = FormatPercent((closedprice - openprice) / openprice, 2)
        
       Else
        Range("L" & Summary_Table_Row).Value = Null
    
      End If

      'Print the total stock volume amount to the Summary Table
      Range("M" & Summary_Table_Row).Value = Total_Stock_volume_Total

       'Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

       'Reset the Stock volume Total
      Total_Stock_volume_Total = 0
      
      openprice = Cells(I + 1, 3).Value
    
End If
  
    Next I

'max value
   Range("R3").Value = WorksheetFunction.Max(Range("L2:L" & lastrow)) * 100
   Dim rowmax As Double
   rowmax = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), (Range("L2:L" & lastrow)), 0)
   Range("Q3").Value = Cells(rowmax + 1, 10).Value
'min value
  Range("R4").Value = WorksheetFunction.Min(Range("L2:L" & lastrow)) * 100
  Dim rowmin As Double
  rowmin = WorksheetFunction.Match(WorksheetFunction.Min(Range("L2:L" & lastrow)), (Range("L2:L" & lastrow)), 0)
  Range("Q4").Value = Cells(rowmin + 1, 10).Value

'max total
  Range("R5").Value = WorksheetFunction.Max(Range("M2:M" & lastrow)) * 100
  Dim rowtotal As Double
  rowtotal = WorksheetFunction.Match(WorksheetFunction.Max(Range("M2:M" & lastrow)), (Range("M2:M" & lastrow)), 0)
  Range("Q5").Value = Cells(rowtotal + 1, 10).Value
 

Next ws

End Sub



