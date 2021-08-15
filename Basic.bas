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
    
    'Set count to 1, this sets the row for the summary table
    count = 1
    countSame = 1
    
    'Find last row of table
    lastRow = Cells(Rows.count, 1).End(xlUp).Row
    
    'Set column headers
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"

    'loop through rows of data for the sheet
    For i = 2 To lastRow

        'checks if current cell and next cell have same ticker value
        If Cells(i + 1, 1).Value <> Cells(i, 1) Then
            'Set ticker value
            ticker = Cells(i, 1).Value
            
             ' Find the close price at latest month

                latestMonthPrice = Cells(i, 6).Value
                
                ' yearly change is dec close - jan open
                yearlyChange = latestMonthPrice - earlyMonthPrice
                
                'Find the percentage change
                If yearlyChange = 0 Then
                    percentageChange = 0 'deals with plnt stock which has all 0 values
                Else
                    percentageChange = (yearlyChange / earlyMonthPrice)
                End If

                'Change percentage change to percenteage format
                Range("K" & count + 1).NumberFormat = "0.00%"
                
                'Change yearly change to number format so cells can be condtionally formatted
                Range("J" & count + 1).NumberFormat = "0.00"
                
                'Add to total stock volume
                totalStockVolume = totalStockVolume + Cells(i, 7).Value
                
                'Print values in summary table
                Range("I" & count + 1).Value = ticker
                Range("J" & count + 1).Value = yearlyChange
                Range("K" & count + 1).Value = percentageChange
                Range("L" & count + 1).Value = totalStockVolume
                
                If yearlyChange < 0 Then ' red if less than 0
                    Range("J" & count + 1).Interior.ColorIndex = 3
                ElseIf yearlyChange > 0 Then ' green if greater than 0
                    Range("J" & count + 1).Interior.ColorIndex = 4
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
                    If Cells(i, 3).Value = 0 Then
                        earlyMonthPrice = Cells(i + 1, 3).Value
                    Else
                        earlyMonthPrice = Cells(i, 3).Value
                        countSame = countSame + 1
                    End If
                End If
                totalStockVolume = totalStockVolume + Cells(i, 7).Value
            End If
        Next i
               
       MsgBox ("Done")
End Sub