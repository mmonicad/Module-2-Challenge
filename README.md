# Module-2-Challenge

Sub Module2()
    'Columns in Worksheet
    'Ticker, Date, Open, High, Low, Close, Vol
    
    Dim ws As Worksheet
    Dim Ticker As String
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim TickerVolume As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim summary_ticker_row As Integer
    Dim LastRow As Long
    
    ' Loop through each worksheet
    For Each ws In Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Add labels for greatest values
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        ' Initialize variables
        summary_ticker_row = 2
        TickerVolume = 0
        open_price = ws.Cells(2, 3).Value
        
        ' Get last row
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Loop through rows
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' New ticker found
                Ticker = ws.Cells(i, 1).Value
                TickerVolume = TickerVolume + ws.Cells(i, 7).Value
                close_price = ws.Cells(i, 6).Value
                
                ' Calculate changes
                yearly_change = close_price - open_price
                If open_price <> 0 Then
                    PercentChange = yearly_change / open_price
                Else
                    PercentChange = 0
                End If
                
                ' Output to summary
                ws.Range("I" & summary_ticker_row).Value = Ticker
                ws.Range("J" & summary_ticker_row).Value = yearly_change
                ws.Range("K" & summary_ticker_row).Value = PercentChange
                ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
                ws.Range("L" & summary_ticker_row).Value = TickerVolume
                
                ' Increment summary row and reset variables
                summary_ticker_row = summary_ticker_row + 1
                TickerVolume = 0
                open_price = ws.Cells(i + 1, 3).Value  ' Reset opening price for next ticker
            Else
                ' Same ticker, accumulate volume
                TickerVolume = TickerVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        'Color code quarterly change
        LastRow = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
        For i = 2 To LastRow
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 10
            Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
    
        
    Next ws
    


End Sub

