Option Explicit

Sub stocks():

    Dim i As Long 'Row number
    Dim cell_vol As LongLong ' Cell volume in column G
    Dim vol_total As LongLong ' Volume total in column L
    Dim ticker As String ' Ticker in column I

    Dim k As Long 'Leaderboard row
    
    Dim ticker_close As Double
    Dim ticker_open As Double
    Dim price_change As Double
    Dim percent_change As Double
    Dim lastRow As Long
    
    Dim ws As Worksheet
    
    Dim greatest_volume As LongLong
    Dim greatest_percent As Double
    Dim lowest_percent As Double
    Dim greatest_volume_ticker As String
    Dim greatest_percent_ticker As String
    Dim lowest_percent_ticker As String
        
    
    ' Loop through every worksheet
    For Each ws In ThisWorkbook.Worksheets
    
    
        ' Get the last row (Bootcamp Spot expert)
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
        vol_total = 0
        k = 2
        
        ' Write leaderboard columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Volume Total"
        
        ' Write second leaderboard columns
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest Percent Change"
        ws.Range("O3").Value = "Lowest Percent Change"
        ws.Range("O4").Value = "Greatest Volume"
        
        ' Assign open for first ticker
        ticker_open = ws.Cells(2, 3).Value
    
        ' Loop rows 2 to 254
        ' Look ahead: Check if next row ticker is different
        ' If ticker is the same, add to the vol_total
        ' If ticker is different, add last row, write out to the leaderboard
        ' Reset the vol_total to 0 for new ticker
            
        For i = 2 To lastRow:
            cell_vol = ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
    
            ' Look ahead
            If (ws.Cells(i + 1, 1).Value <> ticker) Then
                vol_total = vol_total + cell_vol
                
                ' Get the closing price of the ticker
                ticker_close = ws.Cells(i, 6).Value
                price_change = ticker_close - ticker_open
                
                ' Check if open price is 0
                If (ticker_open > 0) Then
                    percent_change = price_change / ticker_open
                Else
                    percent_change = 0
                End If
                
                ' Write to leaderboard
                ws.Cells(k, 9).Value = ticker
                ws.Cells(k, 10).Value = price_change
                ws.Cells(k, 11).Value = percent_change
                ws.Cells(k, 12).Value = vol_total
                
                ' Cell formatting
                If (price_change > 0) Then
                    ws.Cells(k, 10).Interior.ColorIndex = 4 ' Green
                    ws.Cells(k, 11).Interior.ColorIndex = 4 ' Green
                ElseIf (price_change < 0) Then
                    ws.Cells(k, 10).Interior.ColorIndex = 3 ' Red
                    ws.Cells(k, 11).Interior.ColorIndex = 3 ' Red
                Else
                    ws.Cells(k, 10).Interior.ColorIndex = 2 ' White
                    ws.Cells(k, 11).Interior.ColorIndex = 2 ' White
                End If
    
                ' Second leaderboard
                If ticker = ws.Cells(2, 1).Value Then
                    ' Initialize variables
                    greatest_volume = vol_total
                    greatest_percent = percent_change
                    lowest_percent = percent_change
                    greatest_volume_ticker = ticker
                    greatest_percent_ticker = ticker
                    lowest_percent_ticker = ticker
                Else
                    ' Compare variables
                    If vol_total > greatest_volume Then
                        greatest_volume = vol_total
                        greatest_volume_ticker = ticker
                    End If
                    
                   If percent_change > greatest_percent Then
                        greatest_percent = percent_change
                        greatest_percent_ticker = ticker
                    End If

                   If percent_change < lowest_percent Then
                        lowest_percent = percent_change
                        lowest_percent_ticker = ticker
                    End If
                    
                End If
                             
                ' Reset
                vol_total = 0
                k = k + 1
                ' Look ahead, set next ticker open
                ticker_open = ws.Cells(i + 1, 3).Value
            Else
                ' If ticker is not different, add to the volume total
                vol_total = vol_total + cell_vol
            End If
        Next i
        
        ' Format leaderboard cells and columns
        ws.Columns("K:K").NumberFormat = "0.00%"
        ws.Columns("I:L").AutoFit
        ws.Columns("O:O").AutoFit
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Columns("Q").AutoFit
        
        
        ' Write to second leaderboard
        ws.Range("Q2").Value = greatest_percent
        ws.Range("P2").Value = greatest_percent_ticker
        ws.Range("Q3").Value = lowest_percent
        ws.Range("P3").Value = lowest_percent_ticker
        ws.Range("Q4").Value = greatest_volume
        ws.Range("P4").Value = greatest_volume_ticker
        
    Next ws
    
End Sub
