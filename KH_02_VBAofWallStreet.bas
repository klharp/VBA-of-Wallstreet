Attribute VB_Name = "Module1"
'Started this exercise with a single worksheet then modified for multi-year/multi-worksheet

Sub StockAnalysis()

'--Loop through all worksheets--
For Each ws In Worksheets

    'Establish variables
    Dim WorksheetName As String
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim price_change As Double
    Dim percent_change As Double
    Dim volume As Double
    
    
    '--Declare variables--
    open_price = ws.Cells(2, 3).Value 'Starts open price
    close_price = 0
    price_change = 0
    percent_change = 0
    volume = 0
    
    '--Track location--
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2
    
    '--Name headers in worksheet--
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    '--Autofit columns--
    For Each sht In ThisWorkbook.Worksheets
        sht.Cells.EntireColumn.AutoFit
    Next sht
    
    '--Locate last row--
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    '--Loop to find tickers--
    For t = 2 To LastRow
      
        '--Loop each row to determine the tickers--
        If ws.Cells(t, 1).Value <> ws.Cells(t + 1, 1).Value Then
    
            '--Find tickers--
            ticker = ws.Cells(t, 1).Value

            '--Find Close Price--
            close_price = ws.Cells(t, 6).Value
        
            '--Determine Annual Change--
            price_change = (close_price - open_price)
    
            '--Determine percent change and deal with zero fields--
            If open_price = 0 Then
            
                percent_change = 0
            
            Else
        
                percent_change = price_change / open_price
            
            End If
        
            'Debug.Print ticker & " "; open_price & " "; close_price & " "; percent_change
            

        '--Determine Volume--
        volume = volume + ws.Cells(t, 7).Value
            
        '--Display Summary info--
        ws.range("I" & SummaryTableRow).Value = ticker
    
        ws.range("J" & SummaryTableRow).Value = price_change
    
        ws.range("K" & SummaryTableRow).Value = percent_change
        ws.range("K" & SummaryTableRow).NumberFormat = "0.00%"
    
        ws.range("L" & SummaryTableRow).Value = volume
    
        
        '--Reset values--
        SummaryTableRow = SummaryTableRow + 1
    
        volume = 0
        
        open_price = ws.Cells(t + 1, 3).Value 'Starts open price
        
        'close_price = 0
        
        'price_change = 0
        
        'percent_change = 0
         
    Else
    
        '--Add the volume--
        volume = volume + ws.Cells(t, 7).Value
    
    End If
    
    Next t
    
    '--Format colors on summary table--
    LastRowSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For t = 2 To LastRowSummary
    
        If ws.Cells(t, 10).Value > 0 Then
        
            ws.Cells(t, 10).Interior.ColorIndex = 4
            
        Else
            ws.Cells(t, 10).Interior.ColorIndex = 3
            
        End If
    
    Next t
    
    '--Name headers in worksheet--
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 16).Font.Bold = True
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(1, 17).Font.Bold = True
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 15).Font.Bold = True
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 15).Font.Bold = True
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 15).Font.Bold = True
    
    '--Autofit columns--
    For Each sht In ThisWorkbook.Worksheets
        sht.Cells.EntireColumn.AutoFit
    Next sht
    
    LastRowSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For t = 2 To LastRowSummary
    
        '--Find max percent change--
        If ws.Cells(t, 11) = Application.WorksheetFunction.Max(ws.range("K2:K" & LastRowSummary)) Then
            ws.Cells(2, 16).Value = ws.Cells(t, 9).Value
            
            ws.Cells(2, 17).Value = ws.Cells(t, 11).Value
            ws.Cells(2, 17).NumberFormat = "0.00%"
             
        '--Find min percent change--
        ElseIf ws.Cells(t, 11) = Application.WorksheetFunction.Min(ws.range("K2:K" & LastRowSummary)) Then
            
            ws.Cells(3, 16).Value = ws.Cells(t, 9).Value
            
            ws.Cells(3, 17).Value = ws.Cells(t, 11).Value
            ws.Cells(3, 17).NumberFormat = "0.00%"
            
        '--Find max volume--
        ElseIf ws.Cells(t, 12) = Application.WorksheetFunction.Max(ws.range("L2:K" & LastRowSummary)) Then
            ws.Cells(4, 16).Value = ws.Cells(t, 9).Value
            
            ws.Cells(4, 17).Value = ws.Cells(t, 12).Value
            
        End If
        
    Next t
            
Next ws

End Sub

