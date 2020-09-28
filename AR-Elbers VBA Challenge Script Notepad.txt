Attribute VB_Name = "Module1"

Sub VBA_Challange_Final()

 
Dim sheet As Worksheet

For Each sheet In ActiveWorkbook.Worksheets

'clear any previous data in summary table range
sheet.Range("I1:N500").Clear

'Set headers for summary table
sheet.Cells(1, 9).Value = "Ticker Symbol"
sheet.Cells(1, 10).Value = "Ticker Volume"
sheet.Cells(1, 11).Value = "Opening Balance"
sheet.Cells(1, 12).Value = "Closing Balance"
sheet.Cells(1, 13).Value = "Balance Change"
sheet.Cells(1, 14).Value = "Percent Change"
sheet.Range("I1:N1").Font.Bold = True
sheet.Range("I1:N1").Columns.AutoFit
 
    'Set an initial variable for holding the Ticker name
    Dim Ticker_Symbol As String

    'Set an initial variable for holding the total per Ticker Symbol
    Dim Ticker_Total_Volume As Double
    Ticker_Total_Volume = 0
     
    'Keep track of the location for each Ticker Name in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Set variable to hold open and closing values
    Dim Ticker_Open As Double
    Dim Ticker_Close As Double

    'Loop through all Tickers
    For I = 2 To sheet.Cells(Rows.Count, 1).End(xlUp).Row
  
      
    'Check if we are still within the same Ticker Symbol, if it is not...
    If sheet.Cells(I + 1, 1).Value <> sheet.Cells(I, 1).Value Then

    'Set the Ticker Symbol
    Ticker_Symbol = sheet.Cells(I, 1).Value
    'Print the Ticker Symbol in the Summary Table
    sheet.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
      
    'Add to the Ticker Total Volume
    Ticker_Total_Volume = Ticker_Total_Volume + sheet.Cells(I, 7).Value
    'The total stock volume of the stock.
    'Print the Ticker Total Volume to the Summary Table
    sheet.Range("J" & Summary_Table_Row).Value = Ticker_Total_Volume
      
      
    'Find Ticker Close Balance
    Ticker_Close = sheet.Cells(I, 6).Value
    'Print the Ticker Closing Balance to the Summary Table
    sheet.Range("L" & Summary_Table_Row).Value = Ticker_Close
            
      
    'Find Ticker Open Balance
    Ticker_Open = sheet.Cells(I + 1, 3).Value
    'Print the Ticker Open in the Summary Table
    sheet.Range("K" & Summary_Table_Row + 1).Value = Ticker_Open
    sheet.Range("K2") = sheet.Range("C2").Value
      
    'Reset the Ticker Total Volume
    Ticker_Total_Volume = 0
     
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    sheet.Range("M" & Summary_Table_Row).Value = sheet.Range("L" & Summary_Table_Row).Value - sheet.Range("K" & Summary_Table_Row).Value

    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'sheet.Range("N" & Summary_Table_Row).Value = IfError.Range("M" & Summary_Table_Row).Value / Range("K" & Summary_Table_Row).Value
    
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
     If sheet.Range("K" & Summary_Table_Row).Value <> 0 Then
     sheet.Range("N" & Summary_Table_Row).Value = (sheet.Range("M" & Summary_Table_Row).Value / sheet.Range("K" & Summary_Table_Row).Value) * 100
     Else
     sheet.Range("N" & Summary_Table_Row).Value = (sheet.Range("M" & Summary_Table_Row).Value / 1) * 100
     End If
    
    sheet.Range("I" & Summary_Table_Row, "N" & Summary_Table_Row - 1).BorderAround ColorIndex:=1, Weight:=xlThin
    
    'Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
     
    'If the cell immediately following a row is the same Ticker Name...
    Else
      
      'Add to the Ticker Total Volume
      Ticker_Total_Volume = Ticker_Total_Volume + sheet.Cells(I, 7).Value
      
    End If
    
    Next I


            'You should also have conditional formatting that will highlight positive change in green and negative change in red.
            Dim rng As Range
            Dim cell As Range

            Set rng = sheet.Range("M2:N5000")

                For Each cell In rng
                'Cell in red for negative change. If not...
                If cell < 0 Then
                cell.Interior.ColorIndex = 3
                'then cell in green for positive change
                ElseIf cell > 0 Then
                cell.Interior.ColorIndex = 50
                'if no change leave blank... optional extra for myself when testing scripts
                ElseIf cell = 0 Then
                cell.Interior.ColorIndex = 0
                
            End If

            Next



    Next
End Sub



