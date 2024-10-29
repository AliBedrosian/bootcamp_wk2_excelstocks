Attribute VB_Name = "Module11"
Option Explicit

Sub stocks_data()

    Dim total_vol As Double
    Dim i As LongLong
    Dim j As LongLong
    Dim start As LongLong
    Dim row_count As LongLong
    Dim summary_row As LongLong
    Dim ticker_name As String
    Dim open_price As Double
    Dim close_price As Double
    Dim quar_change As Double
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_total_vol As Double
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Quarterly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    
    j = 0
    total_vol = 0
    quar_change = 0
    start = 2
    summary_row = 2
    open_price = Cells(2, 3).Value
    row_count = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For i = 2 To row_count
    
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ticker_name = Cells(i, 1).Value
            
            close_price = Cells(i, 6).Value
            
            quar_change = close_price - open_price
            
            total_vol = total_vol + Cells(i, 7).Value
            
            Range("I" & summary_row).Value = ticker_name
            
            Range("J" & summary_row).Value = quar_change
            
            Range("L" & summary_row).Value = total_vol
            
            Range("K" & summary_row).Value = FormatPercent((close_price - open_price) / open_price)
            
            'conditional formatting
                If (quar_change > 0) Then
                    Range("J" & summary_row).Interior.ColorIndex = 4
                    
                ElseIf (quar_change < 0) Then
                    Range("J" & summary_row).Interior.ColorIndex = 3
                
                End If
            
            'reset
            
            summary_row = summary_row + 1
            open_price = Cells(i + 1, 3).Value
            total_vol = 0
            
        Else
            
            total_vol = total_vol + Cells(i, 7).Value
        
        End If
        
     Next i
    
    Dim greatest_incr As Double
    Dim greatest_decr As Double
    Dim greatest_total As LongLong
    Dim greatest_incr_tick As String
    Dim greatest_decr_tick As String
    Dim greatest_total_tick As String
    
    greatest_incr = Cells(2, 11).Value
    greatest_decr = Cells(2, 11).Value
    greatest_total = Cells(2, 12).Value
    greatest_incr_tick = Cells(2, 9).Value
    greatest_decr_tick = Cells(2, 9).Value
    greatest_total_tick = Cells(2, 9).Value
    
    For j = 2 To summary_row
        If (Cells(j, 11).Value > greatest_incr) Then
            greatest_incr = Cells(j, 11).Value
            greatest_incr_tick = Cells(j, 9).Value
        End If
        
        If (Cells(j, 11).Value < greatest_decr) Then
            greatest_decr = Cells(j, 11).Value
            greatest_decr_tick = Cells(j, 9).Value
        End If
        
        If (Cells(j, 12).Value > greatest_total) Then
            greatest_total = Cells(j, 12).Value
            greatest_total_tick = Cells(j, 9).Value
        End If
        
    Next j
    
    Cells(2, 17).Value = FormatPercent(greatest_incr)
    Cells(3, 17).Value = FormatPercent(greatest_decr)
    Cells(4, 17).Value = greatest_total
    Cells(2, 16).Value = greatest_incr_tick
    Cells(3, 16).Value = greatest_decr_tick
    Cells(4, 16).Value = greatest_total_tick
        
        
    
End Sub
    

