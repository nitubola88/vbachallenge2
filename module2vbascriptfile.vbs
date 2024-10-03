Sub forEachWs()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
 Call ticker1(ws)
 
Call GREATEST(ws)
    Next
End Sub


Sub ticker1(ws)



Dim openval As Double
Dim closingval As Double
Dim ctr As Long
Dim octr As Long
Dim diff As Double
Dim ticker As String
Dim summary_table_row As Long
Dim percetange_change As Double
Dim total_stock As LongLong
summary_table_row = 2
ctr = 0
total_stock = 0
Dim mean As Double
Dim i As Long
total_stock = 0
lastrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To lastrow
    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
octr = (i) - (ctr)
    
      ticker = ws.Cells(i, 1).Value
      openval = ws.Cells(octr, 3).Value
      closingval = ws.Cells(i, 6).Value
     
    diff = closingval - openval
 
    percetange_change = (diff / openval)
   total_stock = total_stock + ws.Cells(i, 7).Value
   
      ws.Range("I" & summary_table_row).Value = ticker

      ws.Range("J" & summary_table_row).Value = diff
      
      
      If (diff < 0) Then

 ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
   
   Else
  ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
    End If
      
     ws.Range("K" & summary_table_row).Value = percetange_change
      ws.Range("K" & summary_table_row).Value = FormatPercent(ws.Range("K" & summary_table_row))
      ws.Range("L" & summary_table_row).Value = total_stock
      
      
      summary_table_row = summary_table_row + 1

  
    
    
      diff = 0
      openval = 0
      closingval = 0
      ctr = 0
      total_stock = 0
    Else

    ctr = ctr + 1
    total_stock = total_stock + ws.Cells(i, 7).Value
    End If

   
     
Next i
ws.Range("i1").Value = "ticker"
ws.Range("j1").Value = "differnce"
ws.Range("k1").Value = "percetange_change"
ws.Range("L1").Value = "total_stock"
ws.Range("O2").Value = "Greatest % increase"
ws.Range("O3").Value = "Greatest % decrease"
ws.Range("O4").Value = "Greatest total volume"
ws.Range("P1").Value = "TICKER"
ws.Range("Q1").Value = "VALUE"


End Sub

Sub GREATEST(ws)
Dim maxval, minval As Double
Dim x As Long
Dim maxstock As LongLong
Dim lastrowperc_change As Long
'Worksheets(ws).Activate
maxval = Application.WorksheetFunction.Max(ws.Range("k:k"))
ws.Range("q2").Value = FormatPercent(maxval)
'MsgBox (maxval)
minval = Application.WorksheetFunction.Min(ws.Range("k:k"))
ws.Range("q3").Value = FormatPercent(minval)

maxstock = Application.WorksheetFunction.Max(ws.Range("l:l"))
ws.Range("q4").Value = maxstock

lastrowperc_change = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

For x = 2 To lastrowperc_change
If ws.Cells(x, 11) = maxval Then
ws.Range("p2").Value = ws.Cells(x, 9)
End If
Next x

For x = 2 To lastrowperc_change
If ws.Cells(x, 11) = minval Then
ws.Range("p3").Value = ws.Cells(x, 9)
End If
Next x

For x = 2 To lastrowperc_change
If ws.Cells(x, 12) = maxstock Then
ws.Range("p4").Value = ws.Cells(x, 9)
End If
Next x
End Sub
