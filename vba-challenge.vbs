' Developed a script that loops all the stocks for one year and output the following information:
' The ticker symbol
' Yearly change from the opening price at the beginning of the given year to the closing price at the end of that year.
' The percent change from the opening price at the beginning of the given year to the closing price at the end of that year.
' The total stock volume of the stock.
' The stock with Greatest % increase, Greatest % decrease and Greatest total volume.
' Conditional formatting that will highlight positive yearly change in green and negative yearly change in red.
' --------------------------------------------------------------------------------------------
Sub Calculatestocks()
'For loop to iterate over all the worksheets
For Each ws In Worksheets
    'Declaring different variables to store the required values
    Dim lastRow As Long
    Dim Ticker As String
    Dim openingprice As Double
    Dim closingprice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim Value As Double
    Dim summarytablerow As Long
    'Finding the last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    summarytablerow = 2
    'Giving the name to columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "YearlyChange"
    ws.Cells(1, 11).Value = "PercentChange"
    ws.Cells(1, 12).Value = "TotalStockVolume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest%Increase"
    ws.Cells(3, 15).Value = "Greatest%Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    'Saving the value of first opening price
    openingprice = ws.Cells(2, 3).Value
    'For loop to find the values of yearly change, percent change and total stock volume
    For i = 2 To lastRow
        'Calculating the total stock value by adding same ticker value to TotalStockVolume variable
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        'If next row is having different ticker value then it will go inside the if condition
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Saving the ticker value
            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & summarytablerow).Value = Ticker
            'Saving the closing price
            closingprice = ws.Cells(i, 6).Value
            'Calculating the yearly change
            YearlyChange = closingprice - openingprice
            ws.Range("J" & summarytablerow).Value = YearlyChange
            'Calculating the percent change
            PercentChange = (closingprice - openingprice) / openingprice
            ws.Range("K" & summarytablerow).Value = PercentChange
            ws.Range("K" & summarytablerow).Style = "Percent"
            ws.Range("K" & summarytablerow).NumberFormat = "0.00%"
            ws.Range("L" & summarytablerow).Value = TotalStockVolume
            'Resetting the total stock volume to 0
            TotalStockVolume = 0
            summarytablerow = summarytablerow + 1
            'Saving the next ticker's opening price
            openingprice = ws.Cells(i + 1, 3).Value
        End If
    Next i
    'Finding the greatest percent increase/decrease and total stock volume with worksheet functions
    ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
    ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
    ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
    ws.Range("Q2,Q3").Style = "Percent"
	ws.Range("Q2,Q3").NumberFormat = "0.00%"
    'For loop to find the ticker with greatest percent increase
    For j = 2 To lastRow
        If (ws.Cells(2, 17).Value = (ws.Cells(j, 11).Value)) Then
            ws.Cells(2, 16).Value = (ws.Cells(j, 9).Value)
        End If
    Next j
    'For loop to find the ticker with greatest percent decrease
    For k = 2 To lastRow
        If (ws.Cells(3, 17).Value = (ws.Cells(k, 11).Value)) Then
            ws.Cells(3, 16).Value = (ws.Cells(k, 9).Value)
        End If
    Next k
    'For loop to find the ticker with greatest total volume
    For l = 2 To lastRow
        If (ws.Cells(4, 17).Value = (ws.Cells(l, 12).Value)) Then
            ws.Cells(4, 16).Value = (ws.Cells(l, 9).Value)
        End If
    Next l
    'For loop to do conditional formatting that will highlight positive yearly change in green and negative yearly change in red.
    For m = 2 To summarytablerow
        If ws.Cells(m, 10).Value >= 0 Then
            ws.Cells(m, 10).Interior.ColorIndex = 4
        Else
           ws.Cells(m, 10).Interior.ColorIndex = 3
        End If
    Next m
Next ws
End Sub