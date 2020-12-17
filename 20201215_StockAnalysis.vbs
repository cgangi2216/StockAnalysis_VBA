Sub CleaResults()
'Loop through each worksheet
    For Each ws In Worksheets
        'Go to worksheet
            Worksheets(ws.Name).Activate
        'Clear cell values & reset cell color
            Range("H:S").Value = ""
            Range("H:S").Interior.ColorIndex = 0
    Next ws
End Sub
Sub CalcChange()
    'https://www.automateexcel.com/vba/list-all-sheets-in-workbook/
    
'Dim variables
    Dim ticker As String
    Dim openprice As Currency
    Dim opendate As Long
    Dim endprice As Currency
    Dim enddate As Long
    Dim volume As Double
    Dim i As Long
    Dim row As Integer
    Dim ws As Worksheet
    
'Greatest % increase
    Dim up_ticker
    Dim up_changerate
'Greatest % decrease
    Dim down_ticker
    Dim down_changerate
'Greatest total volume
    Dim vol_ticker
    Dim vol_changerate

'Loop through each worksheet
    For Each ws In Worksheets
        'Go to worksheet
            Worksheets(ws.Name).Activate
        
        'Column Headers
            Range("j1").Value = "Ticker"
            Range("k1").Value = "Year Change"
            Range("l1").Value = "Percent Change"
            Range("M1").Value = "Total Stock Volume"
            
        'Reset variables
            row = 2
            i = 2
            ticker = Cells(i, 1)
            opendate = Cells(i, 2)
            openprice = Cells(i, 3)
            enddate = Cells(i, 2)
            endprice = Cells(i, 6)
            volume = volume + Cells(i, 7)
    
        'Loop through tickers.
            While Cells(i, 1) <> ""
                'Loop through rows.
                    While Cells(i, 1) = ticker
                        If opendate = 0 Or opendate > Cells(i, 2) Then
                                opendate = Cells(i, 2)
                                openprice = Cells(i, 3)
                            End If
                        If enddate = 0 Or enddate < Cells(i, 2) Then
                                enddate = Cells(i, 2)
                                endprice = Cells(i, 6)
                            End If
                        volume = volume + Cells(i, 7)
                        i = i + 1
                    Wend
                'Write results for row
                    Cells(row, 10).Value = ticker
                    Cells(row, 11).Value = openprice - endprice
                    If openprice - endprice < 0 Then Cells(row, 11).Interior.ColorIndex = 3
                    If openprice - endprice > 0 Then Cells(row, 11).Interior.ColorIndex = 4
                    If openprice = 0 Then Cells(row, 12).Value = 0 Else Cells(row, 12).Value = (openprice - endprice) / openprice
                    If openprice - endprice < 0 Then Cells(row, 12).Interior.ColorIndex = 3
                    If openprice - endprice > 0 Then Cells(row, 12).Interior.ColorIndex = 4
                    Cells(row, 12).NumberFormat = "0.00%"
                    Cells(row, 13).Value = volume
                          
                'Greatest % increase
                    If up_ticker = "" Or up_changerate < Cells(row, 12).Value Then up_ticker = ticker
                    If up_ticker = "" Or up_changerate < Cells(row, 12).Value Then up_changerate = Cells(row, 12).Value
                
                'Greatest % decrease
                    If down_ticker = "" Or down_changerate > Cells(row, 12).Value Then down_ticker = ticker
                    If down_ticker = "" Or down_changerate > Cells(row, 12).Value Then down_changerate = Cells(row, 12).Value
                
                'Greatest total volume
                    If vol_ticker = "" Or vol_changerate < volume Then vol_ticker = ticker
                    If vol_ticker = "" Or vol_changerate < volume Then vol_changerate = volume
                
                'Go to Next Row
                    row = row + 1
                
                'Reset variables
                    ticker = Cells(i, 1)
                    opendate = Cells(i, 2)
                    openprice = Cells(i, 3)
                    enddate = Cells(i, 2)
                    endprice = Cells(i, 6)
                    volume = volume + Cells(i, 7)
            Wend
        'Greatest % increase, % decrease, & total volume
            Range("Q1").Value = "Ticker"
            Range("R1").Value = "Value"
            Range("P2").Value = "Greatest % increase"
            Range("Q2").Value = up_ticker
            Range("R2").Value = up_changerate
            Range("R2").NumberFormat = "0.00%"
            Range("P3").Value = "Greatest % decrease"
            Range("Q3").Value = down_ticker
            Range("R3").Value = down_changerate
            Range("R3").NumberFormat = "0.00%"
            Range("P4").Value = "Greatest total volume"
            Range("Q4").Value = vol_ticker
            Range("R4").Value = vol_changerate
    Next ws
End Sub
