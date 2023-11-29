Set objExcel =CreateObject("Excel.Application")
 objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("https://d.docs.live.net/22a4cee72c1fe257/Desktop/VBA Homework/Multiple-year-stock-data/Multiple_year_stock_data.xlsm")

Sub AnalyzeStockSummary()
    
    ' PART 1
    '-------------------------------------------------------------
    
    ' Declare the worksheet as ws
    
    Dim ws As Worksheet
    
    ' Change the sheet name for each worksheet.
     
    Set ws = ThisWorkbook.Sheets("2018")

    ' Calulate last row of worksheet
        
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Declare variables for data in Summary Array. Initialise lastRow and summaryRow
            
    Dim summaryRow As Long
    summaryRow = 2
    
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim startRow As Long
    Dim i As Long
    
    ' initialising startRow
    
    startRow = 2

    ' Make labels for Summary Array and Bonus Sections
        
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ' PART 2  -  Processing the data set. Condensing da
    ' ------------------------------------------------------------
    
    
    ' Locating changes in Ticker Value, updating closing price and total volume when fount.
        
    openingPrice = ws.Cells(startRow, 3).Value
    totalVolume = 0

    For i = 2 To lastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            closingPrice = ws.Cells(i, 6).Value
            yearlyChange = closingPrice - openingPrice
            
            ' Calculating percent change.
            
            If openingPrice <> 0 Then
               
               percentChange = yearlyChange / openingPrice
              
               
            Else
                percentChange = 0
            
            End If
            
            ' Upating totalVolume if not change in Ticker
                        
            totalVolume = totalVolume + ws.Cells(i, 7).Value

            ' PART 3
            ' ----------------------------------------------------
            
            ' Output results to Summary Array
            
            ws.Cells(summaryRow, 9).Value = ticker
            ws.Cells(summaryRow, 10).Value = yearlyChange
            ws.Cells(summaryRow, 11).Value = percentChange
            ws.Cells(summaryRow, 12).Value = totalVolume

            ' Format as currency and percentage
            
            ws.Cells(summaryRow, 10).NumberFormat = "$#,##0.00"
            ws.Cells(summaryRow, 11).NumberFormat = "#,##0.00%"
            ws.Cells(summaryRow, 12).NumberFormat = "#,##0.00"

            'Test percentage change over the year. Colour Red if negative. Colour Green if positive.
            
            If ws.Cells(summaryRow, 10).Value < 0 Then
                ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
            End If
            
            ' Move to the next summary row
            summaryRow = summaryRow + 1

            ' Reset for next ticker
            If i + 1 <= lastRow Then
                openingPrice = ws.Cells(i + 1, 3).Value
                totalVolume = 0
            End If
        Else
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        End If
    Next i
    
    ' END PARTS 1, 2, AND 3 - Calculation and Display of Summary Array
    
    ' ------------------------------------------------------------
    
    ' PART 4 - BONUS
    
    ' Locate and Display the Tickers with the Greatest % Increase and
    ' Deacrease over the 12 months and most shares traded over the same
    ' period
    
    ' Preliminary declaration of variabls maxvalue and rownummax.
    ' Rownumax is the row in which the maximum value is found.
        
    Dim maxvalue As Double
    Dim rownummax As Long
    maxvalue = ws.Cells(2, 11).Value
    rownummax = 2
    
    Dim minvalue As Double
    Dim rownumin As Long
    minvalue = ws.Cells(2, 11).Value
    rownumin = 2
    
    Dim maxvaluevol As Double
    Dim rownumvol As Long
    maxvaluevol = ws.Cells(2, 12).Value
    rownumvol = 2
    
        ' Locate the maximum value (maxvalue)
        
    For m = 3 To lastRow
        If ws.Cells(m, 11).Value > maxvalue Then
            maxvalue = ws.Cells(m, 11).Value
            rownummax = m
        End If
    Next m

    ' Dispay the values
          
    ws.Cells(2, 16).Value = ws.Cells(4, 9).Value  ' Ticker
    ws.Cells(2, 17).Value = maxvalue              ' % change of Ticker
    ws.Cells(2, 17).NumberFormat = "#,##0.00%"    ' Formating value as a percentage

    ' Locate largest minimum value (minvalue)
    ' Assumed the maximum negative value is the minimum value
    
    For m = 3 To lastRow
        If ws.Cells(m, 11).Value < minvalue Then
            minvalue = ws.Cells(m, 11).Value
            rownumin = m
        End If
    Next m
 
    ' Dispay the minumum values
      
    ws.Cells(3, 16).Value = ws.Cells(rownumin, 9).Value
    ws.Cells(3, 17).Value = minvalue
    ws.Cells(3, 17).NumberFormat = "#,##0.00%"

    ' Locate Greatest Total Volume

    For m = 3 To lastRow
        If ws.Cells(m, 12).Value > maxvaluevol Then
            maxvaluevol = ws.Cells(m, 12).Value
            rownumvol = m
        End If
    Next m
   
   ' Display Gretest Total Volume
      
    ws.Cells(4, 16).Value = ws.Cells(rownumvol, 9).Value
    ws.Cells(4, 17).Value = maxvaluevol
    ws.Cells(4, 17).NumberFormat = "##,####,###,##0"

End Sub

objWorkbook.Close False
objExcel.Quit
Set objWorkbook = Nothing
Set objExcel = Nothing
