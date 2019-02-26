Sub multiple_sheets():
  Dim ws As Worksheet
  Application.ScreenUpdating = False
  For Each ws In Worksheets
      ws.Select
      Call ticker
  Next
  Application.ScreenUpdating = True
End Sub

Sub ticker():
  Dim i As Long
 ' Set an initial variable for holding the ticker
  Dim ticker As String
  ' Set an initial variable that holds the total volume per ticker
  Dim Total_vol As Double
  Total_vol = 0
  ' Keep track of the location for each ticker in the summary table
  Dim Summary_table As Integer
   Summary_table = 2
  
  
   ' Print headers for anticipated results
   Range("I1").Value = "Ticker"
   Range("J1").Value = "Total Stock Volume"
   
 ' Make sure last row is accounted for in each worksheet
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  ' Loop through all ticker
  For i = 2 To lastrow
  ' Check if we are still within the same ticker, if it is not then
   If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    ' Set the ticker
    ticker = Cells(i, 1).Value
    ' Add to the Total volume
    Total_vol = Total_vol + Cells(i, 7).Value
    ' Print the ticker names in the summary table
    Range("I" & Summary_table).Value = ticker
    ' Print the total volume to the summary table
    Range("J" & Summary_table).Value = Total_vol
    ' Add one to thesummary table row
    Summary_table = Summary_table + 1
    ' Reset the total volume
    Total_vol = 0
    ' If the cell immediately following a row is the same ticker
    Else
     ' Add to the total volume
     Total_vol = Total_vol + Cells(i, 7).Value
     
    
   End If
 Next i
 Columns("A:J").AutoFit
 

    
End Sub

