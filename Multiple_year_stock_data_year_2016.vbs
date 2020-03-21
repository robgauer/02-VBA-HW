' 02-VBA-Homework
' Rob Gauer
' Year 2016 Code
'
Sub Ticker_Name_2016()

  ' Set an initial variable for holding the Ticker_Name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total per Ticker_Name
  Dim Ticker_Total As Double
  Ticker_Total = 0

  ' Keep track of the location for each Ticker_Name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all Ticker_Name
  For i = 2 To 705714

    ' Check if we are still within the same Ticker_Name, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker_Name
      Ticker_Name = Cells(i, 1).Value

      ' Add to the Ticker_Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value

      ' Print the Ticker_Name in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Ticker_Name Amount to the Summary Table
      Range("J" & Summary_Table_Row).Value = Ticker_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker_Name Total
      Ticker_Total = 0

    ' If the cell immediately following a row is the same Ticker_Name...
    Else

      ' Add to the Ticker_Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value

    End If

  Next i

End Sub

