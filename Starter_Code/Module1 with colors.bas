Attribute VB_Name = "Module1"
Option Explicit

Sub testing_all_together()

    Dim total_vol As Double
    Dim i As Long
    Dim j As Long
    Dim start As Long
    Dim row_count As Long
    Dim summary_row As Integer
    Dim ticker_name As String
    Dim open_price As Double
    Dim close_price As Double
    Dim quar_change As Double
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_total_vol As Double
    Dim max_row As Integer
    
    
    
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    
    j = 0
    total_vol = 0
    quar_change = 0
    start = 2
    summary_row = 2
    open_price = Cells(2, 3).Value
    row_count = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For i = 2 To row_count
    
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Do stuff
            ticker_name = Cells(i, 1).Value
            
            close_price = Cells(i + 1, 6).Value
            
            quar_change = close_price - open_price
            
            total_vol = total_vol + Cells(i, 7).Value
            
            Range("I" & summary_row).Value = ticker_name
            
            Range("J" & summary_row).Value = quar_change
            
            Range("L" & summary_row).Value = total_vol
            
            Range("K" & summary_row).Value = ((close_price - open_price) / open_price) * 100
            
            'color format
            
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
    
    
    greatest_increase = WorksheetFunction.Max(Range("K:K"))
    greatest_decrease = WorksheetFunction.Min(Range("K:K"))
    greatest_total_vol = WorksheetFunction.Max(Range("K:K"))
    
    Cells(2, 17).Value = greatest_increase
    Cells(3, 17).Value = greatest_decrease
    Cells(4, 17).Value = greatest_total_vol
    
End Sub
    
Sub maxes()

    Dim k As Long
    Dim row_count As Long
    row_count = Cells(Rows.Count, 1).End(xlUp).Row
    
    
   For k = 2 To row_count
        
        If Cells(k, 11).Value = Cells(2, 17).Value Then
            Cells(2, 16).Value = Cells(k, 9).Value
            
        Else
            k = k + 1
                    
        End If
        
    Next k

    
End Sub
