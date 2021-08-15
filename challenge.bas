Sub challenge()
   'Declare variables
    Dim ticker As String ' ticker letter
    Dim yearlyChange As Double ' yearly change (open price at earliest date
                                                 ' of year - close price at latest date of year)
    Dim percentChange As Double ' percent changedyearly change / open price at earliest date
    Dim totalStockVolume As LongLong ' total amount of stock
    Dim count As Integer ' counts the rows of the summary table
    Dim lastRow As Long ' finds the last row of the data
    Dim year As String ' finds the year sold (probably delete)
    Dim month As String '
    Dim earlyMonthPrice As Double 'earliest month stock price
    Dim latestMonthPrice As Double 'latest month stock price
    Dim latestMonth As String 'latest month stock is sold
    Dim earliestMonth As String 'earliest month stock is sold
    Dim countSame As Integer
    Dim greatestPercentage As Double
    Dim greatestPercentageDec As Double
    Dim indexGP As Integer
    Dim indexGD As Integer
    Dim indexGV As Integer
    
    'loop through rows of data for the sheet
    For Each ws In Worksheets
        count = 1
        countSame = 1
        
        lastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        For i = 2 To lastRow

        'checks if current cell and next cell have same ticker value
        If Cells(i + 1, 1).Value <> Cells(i, 1) Then
            'Set ticker value
            ticker = ws.Cells(i, 1).Value
            
             ' Find the close price at latest month

                latestMonthPrice = ws.Cells(i, 6).Value
                
                ' yearly change is dec close - jan open
                yearlyChange = latestMonthPrice - earlyMonthPrice
                
                'Find the percentage change
                If yearlyChange = 0 Then
                    percentageChange = 0 'deals with plnt stock which has all 0 values
                Else
                    percentageChange = (yearlyChange / earlyMonthPrice)
                End If

                'Change percentage change to percenteage format
                ws.Range("K" & count + 1).NumberFormat = "0.00%"
                
                'Change yearly change to number format so cells can be condtionally formatted
                ws.Range("J" & count + 1).NumberFormat = "0.00"
                
                'Add to total stock volume
                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
                
                'Print values in summary table
                ws.Range("I" & count + 1).Value = ticker
                ws.Range("J" & count + 1).Value = yearlyChange
                ws.Range("K" & count + 1).Value = percentageChange
                ws.Range("L" & count + 1).Value = totalStockVolume
                
                If yearlyChange < 0 Then ' red if less than 0
                    ws.Range("J" & count + 1).Interior.ColorIndex = 3
                ElseIf yearlyChange > 0 Then ' green if greater than 0
                    ws.Range("J" & count + 1).Interior.ColorIndex = 4
                ElseIf yearlyChange = 0 Then 'yellow if equal to 0
                Range("J" & count + 1).Interior.ColorIndex = 6
                End If
                
                latestMonthPrice = 0
                
                'Add one to count
                count = count + 1
                countSame = 1
                
                'resets the stock volume to 0 once new stock ticker is found
               totalStockVolume = 0

            Else
                ' Find the open price at earliest month for first stock only
                If countSame = 1 Then
                    If ws.Cells(i, 3).Value = 0 Then
                        earlyMonthPrice = ws.Cells(i + 1, 3).Value
                    Else
                        earlyMonthPrice = ws.Cells(i, 3).Value
                        countSame = countSame + 1
                    End If
                End If
                totalStockVolume = totalStockVolume + Cells(i, 7).Value
            End If
        Next i
        
       'greatest percentage increase
       greatestPercentage = WorksheetFunction.Max(ws.Range("K:K"))
       'finds row number that max value is in and finds the ticker
       indexGP = WorksheetFunction.Match(greatestPercentage, ws.Range("K:K"), 0)
       ws.Range("P2").Value = ws.Range("I" & indexGP)
       ws.Range("Q2").Value = greatestPercentage
       ws.Range("Q2").NumberFormat = "0.00%"
      
      'greatest percentage decrease
       greatestPercentageDec = WorksheetFunction.Min(ws.Range("K:K"))
       'finds row number that min value is in and finds the ticker
       indexGD = WorksheetFunction.Match(greatestPercentageDec, ws.Range("K:K"), 0)
       ws.Range("P3").Value = ws.Range("I" & indexGD)
       ws.Range("Q3").Value = greatestPercentageDec
       ws.Range("Q3").NumberFormat = "0.00%"
       
       'greatest volume
       greatestVol = WorksheetFunction.Max(ws.Range("L:L"))
       'finds row number that max stock value is in and finds the ticker
       indexGV = WorksheetFunction.Match(greatestVol, ws.Range("L:L"), 0)
       ws.Range("P4").Value = ws.Range("I" & indexGV)
       ws.Range("Q4").Value = greatestVol
       
       'adjust column width
       ws.Columns(15).AutoFit
    Next ws
       MsgBox ("Done")
End Sub
