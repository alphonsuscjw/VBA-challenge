Sub stock_market()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
    
        ' Set an initial variable for holding the stock name
        Dim Stock_Name As String
        
        ' Set an initial variable for holding the open price of the year for each stock
        Dim open_price As Double
        
        ' Set an initial variable for holding the close price of the year for each stock
        Dim close_price As Double
        
        ' Set an initial variable for holding the difference between the close price of the year and the open price of the year for each stock
        Dim yearly_change As Double
        
        ' Set an initial variable for holding the percent change in price in one year for each stock
        Dim percent_change As Double
        
        ' Keep track how many rows of raw data there are for each stock
        Dim row_count As Integer
        row_count = 0
    
        ' Set an initial variable for holding the total volume per stock
        Dim Stock_Vol_Total As Double
        Stock_Vol_Total = 0
    
        ' Keep track of the location of each stock in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
      
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Fill in the headings of the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Wrap the text of the headings
        ws.Range("I1:L1").WrapText = True
        
         ' Loop through all the raw stock data
        For I = 2 To lastrow
    
            ' Check if we are still within the same stock, if it is not...
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
    
                  ' Set the stock name
                  Stock_Name = ws.Cells(I, 1).Value
                
                  row_count = row_count + 1
                  
                  ' Set the open price of the year for the stock
                  open_price = ws.Cells(I - row_count + 1, 3).Value
                  
                  ' Set the close price of the year for the stock
                  close_price = ws.Cells(I, 6).Value
                  
                  ' Calculate the yearly change in price for that stock
                  yearly_change = close_price - open_price
                  
                  ' Calculate the percent change in price in that year for that stock
                  percent_change = yearly_change / open_price
            
                  ' Add to the Stock Volume Total for that stock
                  Stock_Vol_Total = Stock_Vol_Total + ws.Cells(I, 7).Value
            
                  ' Print the stock name in to Summary Table
                  ws.Range("I" & Summary_Table_Row).Value = Stock_Name
                  
                  ' Print the yearly change to the Summary Table
                  ws.Range("J" & Summary_Table_Row).Value = yearly_change
                  
                  ' Print the percent change to the Summary Table
                  ws.Range("K" & Summary_Table_Row).Value = percent_change
                  
                  ' Print the stock volume total to the Summary Table
                  ws.Range("L" & Summary_Table_Row).Value = Stock_Vol_Total
            
                  ' Add one to the summary table row
                  Summary_Table_Row = Summary_Table_Row + 1
                  
                  row_count = 0
                  
                  ' Reset the stock volume total for use by the next stock
                  Stock_Vol_Total = 0
    
            ' If the cell immediately following a row is the same stock...
            Else
                  row_count = row_count + 1
                
                  ' Add to the Stock Volume Total for that stock
                  Stock_Vol_Total = Stock_Vol_Total + ws.Cells(I, 7).Value
        
            End If
        
      Next I
      
      ' Conditional formatting for the 'yearly change' and 'percent change' columns
      For I = 2 To Summary_Table_Row - 1
      
            ' Format every value in the 'yearly change' column so that they show 2 decimal points even for values like 0.40
            ws.Cells(I, 10).NumberFormat = "0.00"
            
            ' If the yearly change is positive
            If ws.Cells(I, 10).Value > 0 Then
                  ' Colour the cell green
                  ws.Cells(I, 10).Interior.ColorIndex = 4
            
            ' If the yearly change is negative
            ElseIf ws.Cells(I, 10).Value < 0 Then
                  ' Colour the cell red
                  ws.Cells(I, 10).Interior.ColorIndex = 3
            
            ' If the yearly change is 0
            Else
                  ' Colour the cell yellow
                  ws.Cells(I, 10).Interior.ColorIndex = 6
            End If
            
            ' Format every value in the 'percent change' column so that they are percentage of 2 decimal points
            ws.Cells(I, 11).NumberFormat = "0.00%"
        
      Next I
    
    Next ws
    
End Sub

