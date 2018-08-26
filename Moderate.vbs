 Sub YearlyChange()
    For Each ws In Worksheets
     
     'Add The Yearly Change to the Column
      ws.Range("K1").EntireColumn.Insert
     'Add the word Yearly Change to the column header
      ws.Cells(1,11)= "Yearly Change"
     'Add the Percentage Change to the Column
      ws.Range("L1").EntireColumn.Insert
     'Add the word Percentage change to the column header
      ws.Cells(1,12).Value ="Percentage change"

       Dim close1 As Double
       Dim close2 As Double
       'counter
       Dim Summary_Table_Row As Integer
       Summary_Table_Row = 2
       close1 =ws.Cells(2,6).value
       ' Determine the Last Row
       lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       ' Loop over all stocks

        For i = 2 To lastRow

            If ws.Cells(i,1).value = ws.Cells(i+1,1).value Then
            close1 = close1 + 0
            Else
            close2 = ws.Cells(i, 6).Value
            ws.Cells(Summary_table_Row,11).value = (close2 - close1)
             
            
            If (ws.Cells(Summary_Table_Row,11).Value >) 0 Then 
            ws.Cells(Summary_Table_Row,11).Interior.ColorIndex = 4
            Else
             ws.Cells(Summary_Table_Row,11).Interior.ColorIndex = 3
            End If

            'If close1 = o Then 
              'close1 = 0,0000000001
             'End If

             ws.Cells(Summary_Table_Row,12).value =(close2-close1)/VBA.IIf(close1 = 0, 0.0000000001, close1)
             ws.Cells(Summary_Table_Row,12).style ="percent"
             ws.Cells(Summary_Table_Row,12).numberFormat = "0,00%"

            
              Summary_Table_Row = Summary_Table_Row + 1
              close1 = ws.Cells(i+1,6).Value
            End If
         Next i
      Next ws
          
  End Sub
        
