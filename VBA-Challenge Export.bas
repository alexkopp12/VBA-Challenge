Attribute VB_Name = "Module2"
Sub stockcalc()

Dim ws As Worksheet


For Each ws In Worksheets
    
    ws.Activate
        
    ' Title all collumns
    
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    
    ' Declare all variables
    Dim total_rows As Double
    Dim sum_rows As Double
    Dim row As Double
    Dim ticker_row As Double
    Dim year_start As Double
    Dim year_end As Double
    Dim percentage_change As Double
    Dim volume As Double
    Dim maxinc As Double
    Dim maxdec As Double
    Dim maxvol As Double
    Dim I As Integer
    
    
        
    ' Set initial value for all variables
    ticker_row = 2
    row = 2
    year_start = Cells(2, 3).Value
    year_end = 0
    volume = 0
    total_rows = 0
    maxinc = 0
    maxdec = 0
    maxvol = 0
        
    total_rows = ws.UsedRange.Rows.Count
    

    ' 1. create loop that goes through each row in data set
    For row = 2 To total_rows
        year_end = Cells(row, 6).Value
    ' 2. Verify that ticker symbol is the same
    
        If Cells(row, 1).Value <> Cells(row + 1, 1) Then
            
            volume = Cells(row, 7)
            
            Cells(ticker_row, 10).Value = Cells(row, 1).Value
            
            Cells(ticker_row, 11).Value = year_start - year_end
            
            Cells(ticker_row, 12) = (year_start - year_end) / year_start
            
            Cells(ticker_row, 13) = (volume * year_end)
            
            ticker_row = ticker_row + 1
            year_start = Cells(row + 1, 3)
            
        End If
    
    Next row
     
    
    For I = 2 To ticker_row
           
        Cells(I, 12).Value = FormatPercent(Cells(I, 12).Value, 2)
                
        If Cells(I, 11).Value >= 0 Then
        
            Cells(I, 11).Interior.ColorIndex = 4
            
        Else
        
            Cells(I, 11).Interior.ColorIndex = 3
            
        End If
        
        If maxinc < Cells(I, 12).Value Then
        
            maxinc = Cells(I, 12).Value
            
            Range("P2").Value = Cells(I, 10).Value
            
            Range("Q2").Value = maxinc
            
            Range("Q2").Value = FormatPercent(Range("Q2").Value, 2)
            
        End If
        
        If maxdec > Cells(I, 12).Value Then
         
            maxdec = Cells(I, 12).Value
            
            Range("P3").Value = Cells(I, 10).Value
            
            Range("Q3").Value = maxdec
            
            Range("Q3").Value = FormatPercent(Range("Q3").Value, 2)
            
        End If
        
        If maxvol < Cells(I, 13).Value Then
         
            maxvol = Cells(I, 13).Value
            
            Range("P4").Value = Cells(I, 10).Value
            
            Range("Q4").Value = maxvol
            
        End If
    Next I
    
Next ws


End Sub


