Attribute VB_Name = "Module1"
Sub StockAnalysisFinal()

   'Combined scripts for Stock Analysis and Challenge sections.
		
   'Declare variables and assign values for the stock analysis table; Worksheet,Ticker name,
   'yearly change, percent change, and total stock volume
    
        Dim ws As Worksheet
        
  'loop through all the worksheets
    For Each ws In Worksheets

   
    'Set the variable for holding the ticker name
        Dim Ticker_Name As String
        Ticker_Name = ""
    
    'Set the varable for holding a total stock volume
        Dim TickerVolume As Double
        TickerVolume = 0

    'set the variable to keep track of the location for each ticker name
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
    'Yearly Change is calculated as; (Close Price - Open Price)
    'Percent change is a calculated as ((Close - Open)/Open)*100
        
    ' set the variable to hold the open price and the close price
        Dim open_price As Double
        
    'Set the initial open_price. the succeeding opening prices will be determined by the conditional loop.
        open_price = ws.Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
                 
    
    'Label the Summary Table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

    'Count the number of rows in the first column to track the tickers.
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

     'start the Loop through the rows for each worksheet
        
        For i = 2 To lastrow

            'Check if we are still within the same ticker type, if it is not....
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
             'Set the ticker name
              Ticker_Name = ws.Cells(i, 1).Value

             'Get the total stock volume
              TickerVolume = TickerVolume + ws.Cells(i, 7).Value

              'Print the ticker name in the summary table
              ws.Range("I" & summary_ticker_row).Value = Ticker_Name

              'Print the total stock volume for the ticker in the summary table
              ws.Range("L" & summary_ticker_row).Value = TickerVolume

              'Determine the closing price
              close_price = ws.Cells(i, 6).Value

              'Calculate yearly change
               yearly_change = (close_price - open_price)
              
              'Print the yearly change for each ticker in the summary table
               ws.Range("J" & summary_ticker_row).Value = yearly_change

              'Check for the non-divisibilty condition when calculating the percent change
                If open_price = 0 Then
                    percent_change = 0
                
                Else
                    percent_change = yearly_change / open_price
                
                End If

              'Print the yearly change for each ticker in the summary table
              ws.Range("K" & summary_ticker_row).Value = percent_change
              ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Reset the row counter. Add one to the summary_ticker_row
              summary_ticker_row = summary_ticker_row + 1

              'Reset the total stock volume to zero
              TickerVolume = 0

              'Reset the opening price
              open_price = ws.Cells(i + 1, 3)
            
            Else
              
               'Add the the total stock volume
              TickerVolume = TickerVolume + ws.Cells(i, 7).Value

            End If
        
        Next i

    'Format to highlight positive changes in green and negative changes in red
    'Determine the last row of the summary table to be highlighted

    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code the yearly change with green and red
        For i = 2 To lastrow_summary_table
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
        Next i

' Challenge Summary Output
    'set titles for the summary output

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

    'Determine the maximum and minimum values in column "Percent Change" and
    'the maximum value in column "Total Stock Volume"
    'and assign to the corresponding ticker name and value cells in the summary output
    '
        For i = 2 To lastrow_summary_table
        
            'Determine the maximum percent change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'Determin the minimum percent change
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'Determine the maximum total stock volume
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    
    Next ws
        


End Sub
