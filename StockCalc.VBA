Sub StockCalc()

Dim sh As Worksheet
Dim rw As Range
Dim rowcount As Long
Dim stockrowcount As Long
Dim curr_stock As String
Dim pre_stock As String
Dim curr_volume As Double
Dim total_volume As Double
Dim openingprice As Double
Dim closingprice As Double
Dim yearlychange As Double
Dim pctchange As Double
Dim maxpctchange_stock As String
Dim minpctchange_stock As String
Dim maxvolume_stock As String
Dim maxpctchange As Double
Dim minpctchange As Double
Dim maxvolume As Double
Dim stockrange As Range
Dim stocklastrow As Long

'Set sh = ActiveSheet

'loop through all worksheets in the excel so it will calculate all of them in one run
For Each sh In ActiveWorkbook.Worksheets

rowcount = 0
stockrowcount = 0
'curr_volume = 0
total_volume = 0
maxpctchange = 0
minpctchange = 0
maxvolume = 0

'The worksheet is already sorted by Ticker and Date ascending, if not, we will need sort it first

For Each rw In sh.Rows
  rowcount = rowcount + 1
  If rowcount = 1 Then
    sh.Cells(1, 10) = "Ticker"
    sh.Cells(1, 11) = "Yearly Change"
    sh.Cells(1, 12) = "Percent Change"
    sh.Cells(1, 13) = "Total Stock Volume"
  Else
    'logic for the first record or record with stock ticket change
    If rowcount = 2 Or curr_stock <> sh.Cells(rw.Row, 1).Value Then
      'when stock ticker changes, print out the line for the previous ticker
      If rowcount > 2 Then
        yearlychange = closingprice - openingprice
        pctchange = Round((closingprice - openingprice) * 100 / openingprice, 2)
        sh.Cells(stockrowcount + 1, 10) = curr_stock
        sh.Cells(stockrowcount + 1, 11) = yearlychange
        sh.Cells(stockrowcount + 1, 12) = CStr(pctchange) & "%"
        sh.Cells(stockrowcount + 1, 13) = total_volume
        If maxpctchange < pctchange Then
          maxpctchange = pctchange
          maxpctchange_stock = curr_stock
        End If
        If minpctchange > pctchange Then
          minpctchange = pctchange
          minpctchange_stock = curr_stock
        End If
        If maxvolume < total_volume Then
          maxvolume = total_volume
          maxvolume_stock = curr_stock
        End If
      End If
      If Len(sh.Cells(rw.Row, 1).Value) = 0 Then
        Exit For 'Exit the loop when ticket is empty
      End If
      'when stock ticker changes or it's the first record, reset current stock, opening price and total volume
      curr_stock = sh.Cells(rw.Row, 1).Value
      openingprice = sh.Cells(rw.Row, 3).Value
      total_volume = 0
      'counter for number of unique stock ticker
      stockrowcount = stockrowcount + 1
    End If
    'For all records, we will calculate accumulated total volume and set price as the 'current' closing price
    closingprice = sh.Cells(rw.Row, 3).Value
    curr_volume = sh.Cells(rw.Row, 7).Value
    total_volume = total_volume + curr_volume
  End If
Next rw

  'print out three Calculated Values lines
  sh.Cells(1, 17) = "Ticker"
  sh.Cells(1, 18) = "Value"
  sh.Cells(2, 16) = "Greatest % increase"
  sh.Cells(2, 17) = maxpctchange_stock
  sh.Cells(2, 18) = CStr(maxpctchange) & "%"
  sh.Cells(3, 16) = "Greatest % decrease"
  sh.Cells(3, 17) = minpctchange_stock
  sh.Cells(3, 18) = CStr(minpctchange) & "%"
  sh.Cells(4, 16) = "Greatest total volume"
  sh.Cells(4, 17) = maxvolume_stock
  sh.Cells(4, 18) = maxvolume
  
  'Conditional Formatting
  stocklastrow = sh.Cells(Rows.Count, "K").End(xlUp).Row
  Set stockrange = sh.Range("K2:K" & stocklastrow)
  stockrange.FormatConditions.Delete
  
  stockrange.FormatConditions.Add Type:=xlExpression, Formula1:="=LEN(TRIM(K2))=0"
  stockrange.FormatConditions(1).Interior.Pattern = xlNone
  
  stockrange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
  stockrange.FormatConditions(2).Interior.Color = vbRed 'RGB(255, 0, 0)
  
  stockrange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0"
  stockrange.FormatConditions(3).Interior.Color = vbGreen 'RGB(0, 255, 0)

Next sh

End Sub



