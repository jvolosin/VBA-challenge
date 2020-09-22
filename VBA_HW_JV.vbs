Sub StockMarketTesting()

   'Create a script that will loop through all the stocks for one year and output the following information.
        'Label the Summary Table
          Dim LastRow As Double
          Dim ws As Worksheet
          
          For Each ws In Worksheets
            
            ws.Cells(1, 9) = "Ticker"
            ws.Cells(1, 10) = "Yearly Change"
            ws.Cells(1, 11) = "Percentage Change"
            ws.Cells(1, 12) = "Total Sales Volume"
            ws.Cells(1, 14) = "Total Opening Price"
            ws.Cells(1, 15) = "Total Closing Price"

         
         
        'The ticker symbol
          
          ' Set an initial variable for holding the ticket symbol
  
             Dim ticker As String

          ' Set an initial variable for holding the total stock volume of the stock
  
             Dim total_stock_volume As Double
  
             total_stock_volume = 0
             
             Dim total_opening_price As Double
             Dim total_closing_price As Double
  
             total_opening__price = 0
             total_closing_price = 0


  
          ' Keep track of the location for each ticker symbol in the summary table
  
             Dim Summary_Table_Row As Double
             
             Summary_Table_Row = 2
             
          ' Determine the Last Row
             
             LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

          ' Loop through all ticker symbols
             For i = 2 To LastRow

    
          ' Check if we are still within the same cticker symbol, if it is not:
    
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

          ' Set the ticker symbol
      
             ticker = ws.Cells(i, 1).Value

          ' Add to the Total Stock Volume
      
             total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
             total_opening_price = total_opening_price + ws.Cells(i, 3).Value
             total_closing_price = total_closing_price + ws.Cells(i, 6).Value

          ' Print the ticker symbol in the Summary Table
      
             ws.Range("I" & Summary_Table_Row).Value = ticker

          ' Print the Total Stock Volume to the Summary Table
             ws.Range("L" & Summary_Table_Row).Value = total_stock_volume
             
          ' Print the ticker symbol in the Summary Table
      
             ws.Range("N" & Summary_Table_Row).Value = total_opening_price

          ' Print the Total Stock Volume to the Summary Table
             ws.Range("O" & Summary_Table_Row).Value = total_closing_price
      
          ' Add one to the summary table row
             Summary_Table_Row = Summary_Table_Row + 1
      
          ' Reset the Total Stock Volume
             total_stock_volume = 0
             total_opening__price = 0
             total_closing_price = 0


          ' If the cell immediately following a row is the same symbol:
         Else

          ' Add to the Total Stock Volume
             total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
             total_opening_price = total_opening_price + ws.Cells(i, 3).Value
             total_closing_price = total_closing_price + ws.Cells(i, 6).Value
    End If
        
        
      'Calculate the Yearly Change
        'Create Variable
        
          Dim yearly_change As Double
        
        'Determine the Last Row
        
          LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
          For j = 2 To LastRow

        
          yearly_change = ws.Cells(j + 1, 15).Value - ws.Cells(j + 1, 14).Value
        
          
        'Print the Yearly Change in the Summary Table
          ws.Range("J" & Summary_Table_Row).Value = yearly_change
        
        
        Next j
        
       Next i
       
    Next ws


End Sub
