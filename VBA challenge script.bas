Attribute VB_Name = "Module1"
Sub StockChanges()
Attribute StockChanges.VB_ProcData.VB_Invoke_Func = " \n14"
'
' set ws variable
Dim ws As Worksheet

' loop through worksheets
For Each ws In Worksheets


' Set variables
  Dim Ticker_Name As String

  Dim Ticker_Total As Double
  Ticker_Total = 0

  Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
  Dim LastRow As Double
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  Dim open_price As Double
  Dim close_price As Double
  Dim yearly_change As Double
  Dim percent_yearly_change As Double
 
 

  open_price = ws.Cells(2, 3).Value
  
  ' Loop through data
  
  For i = 2 To LastRow
   
   
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
      Ticker_Name = ws.Cells(i, 1).Value
      Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
     
      close_price = ws.Cells(i, 6).Value
     
      yearly_change = close_price - open_price
     
      percent_yearly_change = (yearly_change / open_price)
     
      open_price = ws.Cells(i + 1, 3).Value

      ' Print in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

      ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
     
      ws.Range("J" & Summary_Table_Row).Value = yearly_change
     
      ws.Range("K" & Summary_Table_Row).Value = percent_yearly_change

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
     
      ' Reset
      Ticker_Total = 0

   
    Else

      ' Add to the Brand Total
      Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

    End If

  Next i
  
' add header to columns
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Yearly Change"
ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
ws.Range("L1").Value = "Total Stock Volume"

' loop through summary table
For i = 2 To Summary_Table_Row

' set conditions
 If ws.Cells(i, 10).Value < 0 Then
  ws.Cells(i, 10).Interior.ColorIndex = 3
 Else
  ws.Cells(i, 10).Interior.ColorIndex = 4
 End If
 

Next i
    


' add functionality
ws.Cells(2, 15).Value = "Greatest % increase"
ws.Cells(3, 15).Value = "Greatest % decrease"
ws.Cells(4, 15).Value = "Greatest total value"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

Next ws


End Sub



