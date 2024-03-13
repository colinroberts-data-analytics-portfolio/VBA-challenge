Attribute VB_Name = "Module1"
Sub Market_Analysis_Challenge_2()

    ' Declare Variables------------------------------------------
    
    Dim ws As Worksheet
    Dim ticker As String
    Dim lastRow As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Integer
    
    
    ' Loop --------------------------------------------------------------------
    ' Loop through worksheets        https://stackoverflow.com/questions/21918166/excel-vba-for-each-worksheet-loop
    For Each ws In ThisWorkbook.Sheets
    
        ' Start variables  https://stackoverflow.com/questions/27065840/meaning-of-cells-rows-count-a-endxlup-row
        ' Iterate and loop through each worksheet / set the summary row /  determine last row with symbols in ticker / initialize total volume variable.
        summaryRow = 2
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        totalVolume = 0
        
        ' Set headers for summary table  https://stackoverflow.com/questions/62975110/vba-script-to-format-cells-within-a-column-range-only-formats-the-first-sheet-in
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Loop through each row in the sheet
        For i = 2 To lastRow
            ' Check if the current row has a different ticker symbol  https://www.bing.com/search?q=%27+Check+if+the+current+row+has+a+different+ticker+symbol+++++++++++++If+ws.Cells%28i+%2B+1%2C+1%29.Value+%3C%3E+ws.Cells%28i%2C+1%29.Value+Then+++++++++++++++++%27+Set+ticker+symbol+++++++++++++++++ticker+%3D+ws.Cells%28i%2C+1%29.Value+++++++++++++++++&form=ANNTH1&refig=ee330c174ca24736a0e455c4c0322639&pc=U531
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set----------------------------------------
                ' Realize Ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                ' Realize Closing price              https://stackoverflow.com/questions/76548179/dont-know-how-to-fix
                closingPrice = ws.Cells(i, 6).Value
                
                ' Realize Yearly change
                yearlyChange = closingPrice - openingPrice
                
                ' Realize Percent change       https://money.stackexchange.com/questions/84534/what-is-the-correct-answer-for-percent-change-when-the-start-amount-is-zero-doll
                If openingPrice <> 0 Then
                    percentChange = (yearlyChange / openingPrice) * 100
                Else
                    percentChange = 0
                End If
                
                ' Output----------------------------------------
                ' Add values to summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Green or Red cells ----------------------------------------
                ' Realize cells determined on changes yearly
                If yearlyChange >= 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                End If
                
                ' Clear variables and check for nxt ticker        https://stackoverflow.com/questions/42980386/how-to-reset-variables-or-declarations-vba
                openingPrice = ws.Cells(i + 1, 3).Value
                totalVolume = 0
                summaryRow = summaryRow + 1
            Else
                ' Collect and add total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Realize the row with the greatest percent increase / decrease / total volume
        FindGreatestValues ws
    Next ws
End Sub

Sub FindGreatestValues(ws As Worksheet)

    ' Declare Variables------------------------------------------
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxTickerIncrease As String
    Dim maxTickerDecrease As String
    Dim maxTickerVolume As String
    Dim lastRow As Integer
    
    lastRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    
    'Set Values -------------------------------------------------------------------------
    ' Set max values                  https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.cells
    maxIncrease = ws.Cells(2, 11).Value
    maxDecrease = ws.Cells(2, 11).Value
    maxVolume = ws.Cells(2, 12).Value
    maxTickerIncrease = ws.Cells(2, 9).Value
    maxTickerDecrease = ws.Cells(2, 9).Value
    maxTickerVolume = ws.Cells(2, 9).Value
    
    ' Loop ----------------------------------------------------------------------------
    ' Loop through summary table to find max values             https://stackoverflow.com/questions/45072650/finding-max-value-of-a-loop-with-vba
    For i = 2 To lastRow
        ' Realize greatest percent increase
        If ws.Cells(i, 11).Value > maxIncrease Then
            maxIncrease = ws.Cells(i, 11).Value
            maxTickerIncrease = ws.Cells(i, 9).Value
        End If
        
        ' Realize greatest percent decrease              https://www.mashupmath.com/blog/calculating-percent-decrease
        If ws.Cells(i, 11).Value < maxDecrease Then
            maxDecrease = ws.Cells(i, 11).Value
            maxTickerDecrease = ws.Cells(i, 9).Value
        End If
        
        ' Realize greatest total volume                https://www.exceldome.com/solutions/if-a-cell-is-greater-than-a-specific-value/
        If ws.Cells(i, 12).Value > maxVolume Then
            maxVolume = ws.Cells(i, 12).Value
            maxTickerVolume = ws.Cells(i, 9).Value
        End If
    Next i
        ' Set----------------------------------------
        ' Set headers for ticker and value
    ws.Cells(1, 15).Value = " "
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ' Output----------------------------------------
    ' Output results for greatest percent increase, decrease, and total volume       https://www.exceldome.com/solutions/if-a-cell-is-greater-than-a-specific-value/
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(2, 16).Value = maxTickerIncrease
    ws.Cells(3, 16).Value = maxTickerDecrease
    ws.Cells(4, 16).Value = maxTickerVolume
    ws.Cells(2, 17).Value = maxIncrease & "%"
    ws.Cells(3, 17).Value = maxDecrease & "%"
    ws.Cells(4, 17).Value = maxVolume





End Sub
