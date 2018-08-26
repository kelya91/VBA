Sub CreatestChange()
 For each ws in Worksheets
   Greatest_increase = 0
   Greatest_decrease = 0
   Greatest_Total_Volume = 0

   Dim Ticker1 As String
   Dim Ticker2 As String
   Dim Ticker3 As String

'Fill Headers
   ws.Cells(2,15).value ="Greatest%Increase"
   ws.Cells(3,15).value ="Greatest%Decrease"
   ws.Cells(4,15).value = "Greatest total volume"
   ws.Cells(1,16).value= "Ticker"
   ws.Cells(1,17).value ="Value"


'Determine Last row
  LastRow = ws.Cells(Rows.Count,9).End(xlUP).Row
'Loop all stocks in all sheets
  For i = 2 to LastRow
     If ws.Cells(i,12).Value > Greatest_increase Then
      Greatest_increase = ws.Cells(i,12).Value
      Ticker1 = ws.Cells(i,9).Value
     End If

     If ws.Cells(i,12).Value < Greatest_decrease Then
       Greatest_decrease = ws.Cells(i,12).value
       Ticker2 = ws.Cells(i,9).value
     End If

     If ws.Cells(i,10).value > Greatest_Total_Volume Then
        Greatest_Total_Volume = ws.Cells(i,10).value
        Ticker3 = ws.Cells(i,9).value
     End If
  Next i
   
   'Write values
     ws.Cells(2,16).Value = Ticker1
     ws.Cells(3,16).value =Ticker2
     ws.Cells(4,16).value =Ticker3

     ws.Cells(2,17).value = Greatest_increase
       ws.Cells(2,17).style ="percent"
       ws.Cells(2,17).NumberFormat= "0.00%"
     ws.Cells(3,17).value =Greatest_decrease
       ws.Cells(3,17).style ="percent"  

       ws.Cells(3,17).NumberFormat= "0.00%"
     ws.Cells(4,17).value = Greatest_Total_Volume
 
  Next ws

End Sub