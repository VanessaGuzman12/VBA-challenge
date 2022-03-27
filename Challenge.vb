Sub Greatest()

Dim Max_percent As Double
Dim Min_percent As Double
Dim Total_volume As Double


Dim ticker_max As String
Dim ticker_min As String
Dim ticker_total As String

Dim column As Integer
Dim column2 As Integer

Total_volume = Cells(2, 12).Value

Min_percent = Cells(2, 11).Value

Max_percent = Cells(2, 11).Value

column = 11

column2 = 12

lastrow = Cells(Rows.Count, 9).End(xlUp).Row


    For i = 2 To lastrow
    
    
        If Cells(i, column).Value < Min_percent Then
        Min_percent = Cells(i, column)
        ticker_min = Cells(i, 9).Value
         
        End If
 
 
        If Cells(i, column).Value > Max_percent Then
        Max_percent = Cells(i, column)
        ticker_max = Cells(i, 9).Value
        
          
        End If
        
        If Cells(i, column2).Value > Total_volume Then
        Total_volume = Cells(i, column2)
        total_ticker = Cells(i, 9).Value
        
        
        End If
        
        
        
    Next i
    
    
Range("Q3") = Min_percent
Range("Q2") = Max_percent
Range("Q4") = Total_volume

Range("P3") = ticker_min
Range("P2") = ticker_max
Range("P4") = total_ticker

Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = "Greatest Total Volume"
Range("P1") = "Ticker"
Range("Q1") = "Value"

Range("Q3,Q2").NumberFormat = "0.00%"


End Sub