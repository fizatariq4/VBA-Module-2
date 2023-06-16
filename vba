Sub stocks()

'set your variables here

  Dim ws As Worksheet
  
  Dim lastrow As Long
  
  Dim i As Long
  
  Dim cellValue As Variant
  
  Dim ticker_name As String
  
  Dim ticker_vol As Double
  
  ticker_vol = 0
  Dim ticker_sum As Integer
  ticker_sum = 2
  
  Dim openprice As Double
  openprice = Cells(2, 3).Value
  
  Dim closeprice As Double
  
  Dim perchange As Double

'Make sure the functions run in each sheet

  For Each ws In Worksheets
  
  'define your values
  
        openprice = ws.Cells(2, 3).Value
        'ticker volume
        ticker_vol = 0
        'ticker summary
        ticker_sum = 2
        
        'name your cells
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
'First loop
        For i = 2 To lastrow
'Print the ticker names and their volume
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker_name = ws.Cells(i, 1).Value
                ticker_vol = ticker_vol + ws.Cells(i, 7).Value
'Calculate the yearly change using the open price and the close price
                yearchange = (ws.Cells(i, 6) - openprice)

                If openprice = 0 Then
                    perchange = 0
                Else
                    perchange = yearchange / openprice
                End If
'Start the summary table's values and format them
                ws.Range("I" & ticker_sum).Value = ticker_name
                ws.Range("J" & ticker_sum).Value = yearchange
                ws.Range("j" & ticker_sum).NumberFormat = "0.00"
                ws.Range("K" & ticker_sum).Value = perchange
                ws.Range("K" & ticker_sum).NumberFormat = "0.00%"
                ws.Range("L" & ticker_sum).Value = ticker_vol
   
                ticker_sum = ticker_sum + 1
                ticker_vol = 0
                openprice = ws.Cells(i + 1, 3)

            Else
                ticker_vol = ticker_vol + ws.Cells(i, 7).Value

            End if 
'Conditional formatting
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            End If
            If ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
            
        
        Next i
        
'Second loop through summary table to find greatest increase, reatest decrease, and greatest total volume
        For i = 2 To ticker_sum
        
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & ticker_sum)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            End If
            
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & ticker_sum)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            End If
            
            If ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & ticker_sum)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
            
        Next i
 Next ws
End Sub
