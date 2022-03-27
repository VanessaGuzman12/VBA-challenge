Sub Ticker()
 
 Dim Ticker As String
 
 Dim ticker_total As Integer
 Dim Summary_Table_Row As Integer
 Dim Days_counter As Integer
 
 Dim Opening As Double
 Dim Closing As Double
 Dim Yearly_change As Double
 Dim Percent_change As Double
 Dim volume_total As Double
 
    Days_counter = 0
    Summary_Table_Row = 2
    ticker_total = 0
  
 lastrow = Cells(Rows.Count, 1).End(xlUp).Row
 
  For i = 2 To lastrow

    
    Days_counter = Days_counter + 1
    volume_total = volume_total + Cells(i, 7).Value
    
    
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker = Cells(i, 1).Value
    ticker_total = ticker_total + 1
    Range("I" & Summary_Table_Row).Value = Ticker
  
    Opening = Cells(i - Days_counter + 1, 3).Value
      
    Closing = Cells(i, 6).Value
    
    Yearly_change = (Closing - Opening)
    Range("J" & Summary_Table_Row).Value = Yearly_change
    
   If Opening = 0 Then
   Percent_change = 0
   
   Else
   Percent_change = (Closing - Opening) / (Opening)
   End If
   
   
    Range("k" & Summary_Table_Row).Value = Percent_change
    
    Range("l" & Summary_Table_Row).Value = volume_total
    
        
        
        If Yearly_change >= 0 Then
       Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      
       Else
       Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
       
       End If
       
    Days_counter = 0
    volume_total = 0
    Summary_Table_Row = Summary_Table_Row + 1
 
    
   End If

   
   Cells(i, 11).NumberFormat = "0.00%"

  Next i

Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"




End Sub

