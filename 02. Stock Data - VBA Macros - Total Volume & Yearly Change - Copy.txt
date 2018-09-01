Attribute VB_Name = "Module1"
Sub StockVolume()
' Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
' You will also need to display the ticker symbol to coincide with the total volume.

Dim ws As Worksheet

For Each ws In Worksheets

' Set an initial variable for holding the stock name
Dim stockname As String

' Set an initial variable for holding the total per stock
Dim stocktotal As Double
stocktotal = 0

' Keep track of the location for each credit card brand in the summary table
Dim SummaryTableRow As Integer
SummaryTableRow = 2
  
' Count the number of rows
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Loop through all stocks
For i = 2 To lastrow
     ' Establish year open value
    yearopen = 0
    ' If the stock above is the same but the stock below is different, yearclose is the close value of that row.
    yearclose = 0

    ' Check if we are still within the same stock, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    ' Set the stock name
    stockname = ws.Cells(i, 1).Value

    ' Add to the stock total
    stocktotal = stocktotal + ws.Cells(i, 7).Value

    ' Print the Stock Name in the Summary Table
      ws.Range("J" & SummaryTableRow).Value = stockname

    ' Print the Stock Total to the Summary Table
     ws.Range("K" & SummaryTableRow).Value = stocktotal
     

    ' If loop encounters stock without year open value, establish year open value of that row
    If yearopen = 0 Then
    yearopen = ws.Cells(i, 3).Value
    Else: yearopen = yearopen
    End If


    If ws.Cells(i - 1, 1) = ws.Cells(i, 1) And ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    yearclose = ws.Cells(i, 6).Value
    Else: yearclose = yearclose
    End If

    ' Calculate yearly change
    yearlychange = yearclose - yearopen
    
    ' Print the Yearly Change to the Summary Table
     ws.Range("L" & SummaryTableRow).Value = yearlychange

    If ws.Cells(i, 12).Value < 0 Then
    ws.Cells(i, 12).Interior.ColorIndex = 3
    Else: ws.Cells(i, 12).Interior.ColorIndex = 4
    End If

    ' Add one to the summary table row
    SummaryTableRow = SummaryTableRow + 1
      
    ' Reset the Brand Total
    stocktotal = 0
    
     ' Reset year open value
    yearopen = 0
    
    ' Reset year close value
    yearclose = 0
        
    ' Reset yearlychange value
    yearlychange = 0

    ' If the cell immediately following a row is the same stock...
    Else

    ' Add to the Stock Total
      stocktotal = stocktotal + ws.Cells(i, 7).Value

    End If

  Next i

  
Next
  
End Sub
Sub Reset()

Dim ws As Worksheet

For Each ws In Worksheets

Range("J2:M5000").Clear

Next


End Sub
