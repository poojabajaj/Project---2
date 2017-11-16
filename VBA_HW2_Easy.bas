Attribute VB_Name = "Module1"
 Sub Sum_Ticker()
 
    ' ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
    
  ' Set an initial variable for holding the ticker
    Dim ticker As String
  ' Set an initial variable for holding the total for each ticker
    Dim Total_Volume As Double
    Total_Volume = 0
  ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  ' Loop through all tickers
    For i = 2 To LastRow
    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ' Set the ticker
        ticker = ws.Cells(i, 1).Value
    ' Add to the Total volume
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    ' Print the ticker in the Summary Table
     ws.Range("I" & Summary_Table_Row).Value = ticker
       'Range("G" & Summary_Table_Row).Value = Brand_Name
      ' Print the Brand Amount to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Total_Volume
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total volume
      Total_Volume = 0
    ' If the cell immediately following a row is the same ticker...
    Else
      ' Add to the Total volume
      Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    End If
  Next i
   Next ws
End Sub
