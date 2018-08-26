Sub StockData()
    For Each ws In Worksheets

          ws.Cells(1, 9).Value = "Ticker"
          ws.Cells(1, 10).Value = "Total stock volume"
         

        ' Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        
        WorksheetName = ws.Name
        ' MsgBox WorksheetName
        ' Set an initial variable for holding the ticker
         Dim Ticker As String
     
         ' Set an initial variable for holding the total stock volume per ticker
          Dim Ticker_total As Double
              Ticker_total = 0
            
          ' Keep track of the location for each ticker in the summary table
          Dim Summary_Table_Row As Integer
          Summary_Table_Row = 2
          
          ' Loop through all tickers
          For i = 2 To lastRow
        
            ' Check if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              ' Set the ticker
              Ticker = ws.Cells(i, 1).Value
        
              ' Add to the  Ticker_Total
              Ticker_total = Ticker_total + ws.Cells(i, 7).Value
        
              ' Print the ticker in the Summary Table
              ws.Range("I" & Summary_Table_Row).Value = Ticker
        
              ' Print Total_stock to the Summary Table
              ws.Range("J" & Summary_Table_Row).Value = Ticker_total

              ' Add one to the summary table row
              Summary_Table_Row = Summary_Table_Row + 1
              
              ' Reset the Ticker_Total
              Ticker_total = 0
        
            ' If the cell immediately following a row is the same ticker...
            Else
        
              ' Add to the Ticker_Total
              Ticker_total = Ticker_total + ws.Cells(i, 7).Value
        
            End If
        
          Next i
          
        Next ws
  End Sub