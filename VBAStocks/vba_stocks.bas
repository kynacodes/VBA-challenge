Attribute VB_Name = "Module1"
'Run on all worksheets
Sub all_sheets()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call stock_ticker
    Next
    Application.ScreenUpdating = True
End Sub

Sub stock_ticker()

'Add column headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    
'Set variables
    Dim ticker As String
    Dim opening_price As Double
        opening_price = 0
    Dim closing_price As Double
        closing_price = 0
    Dim price_change As Double
        price_change = 0
    Dim opening_percentage As Double
        opening_percentage = 0
    Dim closing_percentage As Double
        closing_percentage = 0
    Dim percent_change As Double
        percent_change = 0
    Dim total_volume As Double
        total_volume = 0
    Dim table_row_tracker As Integer
        table_row_tracker = 2
    Dim max As Double
        max = 0
    Dim min As Double
        min = 0
    Dim max_total_volume As Double
        max_total_volume = 0
    
    'Loop through table
    For i = 2 To 70926
       
       'Search for opening date in order to get opening price
       If Right(Cells(i, 2), 4) = "0101" Then
            opening_price = Cells(i, 3)
       End If
       
       total_volume = total_volume + Cells(i, 7).Value
       
       'Checks to see if cell belongs to same ticker
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       
            'Print each ticker entry once to table
            ticker = (Cells(i, 1).Value)
            Range("I" & table_row_tracker).Value = ticker
            
           'Calculate closing price
            closing_price = Cells(i, 6).Value
            price_change = closing_price - opening_price
            Cells(table_row_tracker, 10).Value = price_change
             
            'Conditional Formatting
            If price_change >= 0 Then
                Cells(table_row_tracker, 10).Interior.Color = vbGreen
                Else: Cells(table_row_tracker, 10).Interior.Color = vbRed
            End If
            
            'Calculate percent change
            If price_change <> 0 Then
               percent_change = price_change / opening_price
            End If
            
            Cells(table_row_tracker, 11).Value = percent_change
            
            'Conditional Formatting
            Cells(table_row_tracker, 11).NumberFormat = "0.00%"
            If percent_change >= 0 Then
                Cells(table_row_tracker, 11).Interior.Color = vbGreen
                Else: Cells(table_row_tracker, 11).Interior.Color = vbRed
            End If
            
            'Calculate total stock volume
            Cells(table_row_tracker, 12).Value = total_volume
            
            'Create tag for greatest percent increase
            If percent_change > max Then
                max = percent_change
                tag_max = Cells(table_row_tracker, 9).Value
            End If
        
            'Populate and format increase
            Cells(2, 15).Value = tag_max
            Cells(2, 16).Value = max
            Cells(2, 16).NumberFormat = "0.00%"
        
            'Create tag for decrease
            If percent_change < min Then
                min = percent_change
                tag_min = Cells(table_row_tracker, 9).Value
            End If
        
            'Populate and format decrease
            Cells(3, 15).Value = tag_min
            Cells(3, 16).Value = min
            Cells(3, 16).NumberFormat = "0.00%"
        
            'Create tag for max total stock volume
            If total_volume > max_total_volume Then
                max_total_volume = total_volume
                tag_max_total_volume = Cells(table_row_tracker, 9).Value
            End If
        
            Cells(4, 15).Value = tag_max_total_volume
            Cells(4, 16).Value = max_total_volume
        
            
    'Advances to next row in table
    table_row_tracker = table_row_tracker + 1
            
    'Resets volume to zero after adding one ticker's contents
    total_volume = 0
        
       End If

    Next i

End Sub
