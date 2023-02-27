Attribute VB_Name = "Module1"
Sub addHeadersAndFormat()

    'List the relevant headers for the ticker-by-ticker analysis
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Format the columns containing the ticker-by-ticker analysis
    Range(Columns(9), Columns(12)).HorizontalAlignment = xlCenter
    Columns(9).ColumnWidth = 12.5
    Range(Columns(10), Columns(11)).ColumnWidth = 15
    Columns(12).ColumnWidth = 18.5
    
    'List the relevant headers for the max-min analysis
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    'Format the columns containing the max-min analysis
    With Columns(15)
    
        .ColumnWidth = 18.5
        .HorizontalAlignment = xlRight
        
    End With
    
    With Range(Columns(16), Columns(17))
    
        .HorizontalAlignment = xlCenter
    
    End With
    
    Columns(16).ColumnWidth = 12.5
    Columns(17).ColumnWidth = 18.5

End Sub

Sub evaluateTickers()

    'Declare variables for the ticker-by-ticker analysis
    Dim ticker_count As Long
    Dim row_count As Long
    Dim opening_price As Double
    Dim closing_price As Double
    Dim volume As Double
    
    'Retrieve the number of rows, which will determine the count of the for-loop in the ticker-by-ticker analysis
    Cells(1, 1).Select
    Selection.End(xlDown).Select
    row_count = Selection.Row
    
    'Initial value for the ticker count
    ticker_count = 0
    
    'Loop through the raw data for the ticker-by-ticker analysis
    For i = 2 To row_count
    
        'Conditional for the first day of the year (i.e., when mmdd is 0102)
        If Right(Cells(i, 2).Value, 4) = "0102" Then
        
            'If mmdd is 0102, we know we're looking at a new ticker
            'Increment ticker count and print ticker in column I
            ticker_count = ticker_count + 1
            Cells(ticker_count + 1, 9).Value = Cells(i, 1).Value
            
            'Opening price is defined as the value in the <open> column when mmdd is 0102
            opening_price = Cells(i, 3).Value
            
            'Store volume on opening day
            volume = Cells(i, 7).Value
        
        'Conditional for the last day of the year (i.e., when mmdd is 1231)
        ElseIf Right(Cells(i, 2).Value, 4) = "1231" Then
        
            'Closing price is defined as the value in the <close> column when mmdd is 1231
            closing_price = Cells(i, 6).Value
            
            'Compute yearly & percent change based on opening & closing price
            'Print yearly & percent change in columns J & K, respectively
            Cells(ticker_count + 1, 10).Value = closing_price - opening_price
            Cells(ticker_count + 1, 11).Value = (closing_price - opening_price) / opening_price
            
            'Increase total volume by the volume on the closing day
            'Print total volume in column L
            volume = volume + Cells(i, 7).Value
            Cells(ticker_count + 1, 12).Value = volume
        
        Else
        
            'Increase total volume by the daily volume amounts on every other day of the year
            volume = volume + Cells(i, 7).Value
        
        End If
    
    Next i
    
    'Declare variables for max-min analysis
    Dim max_pct_change As Double
    Dim min_pct_change As Double
    Dim max_vol As Double
    Dim max_pct_ticker As String
    Dim min_pct_ticker As String
    Dim max_vol_ticker As String

    'Initial values for the max-min analysis
    max_pct_change = 0
    min_pct_change = 0
    max_vol = 0
    max_pct_ticker = ""
    min_pct_ticker = ""
    max_vol_ticker = ""
    
    'Loop through the ticker-by-ticker analysis
    For j = 2 To ticker_count + 1
    
        'Conditional to catch percent changes that exceed the previous max
        If Cells(j, 11).Value > max_pct_change Then
        
            'Store the new max percent change & corresponding ticker
            max_pct_change = Cells(j, 11).Value
            max_pct_ticker = Cells(j, 9).Value
        
        'Conditional to catch percent changes that fall below the previous min
        ElseIf Cells(j, 11).Value < min_pct_change Then
        
            'Store the new min percent change & corresponding ticker
            min_pct_change = Cells(j, 11).Value
            min_pct_ticker = Cells(j, 9).Value
        
        End If
        
        'Conditional to catch total volumes that exceed the previous max
        If Cells(j, 12).Value > max_vol Then
        
            'Store the new max volume & corresponding ticker
            max_vol = Cells(j, 12).Value
            max_vol_ticker = Cells(j, 9).Value
        
        End If
    
    Next j
    
    'Print the results of the max-min analysis
    Cells(2, 16).Value = max_pct_ticker
    Cells(2, 17).Value = max_pct_change
    Cells(3, 16).Value = min_pct_ticker
    Cells(3, 17).Value = min_pct_change
    Cells(4, 16).Value = max_vol_ticker
    Cells(4, 17).Value = max_vol
    
    'Return to the top of the sheet
    Cells(1, 1).Select
    
End Sub


Sub formatYearlyChange()

    'Select all rows containing data in column J
    Cells(2, 10).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    'Apply conditional formatting to column J
    With Selection
    
        .FormatConditions.Add Type:=xlExpression, Formula1:="=J2<0"
        .FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
        End With
        
        .FormatConditions(1).StopIfTrue = False
        .FormatConditions.Add Type:=xlExpression, Formula1:="=J2>0"
        .FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 5296274
            .TintAndShade = 0
        End With
    
        .FormatConditions(1).StopIfTrue = False
    
    End With

End Sub

Sub formatPercentChange()
    
    'Select all rows containing data in column K
    Cells(2, 11).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    'Apply percent formatting and conditional formatting to column K
    With Selection
    
        .NumberFormat = "0.00%"
        
        .FormatConditions.Add Type:=xlExpression, Formula1:="=K2<0"
        .FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
        End With
        
        .FormatConditions(1).StopIfTrue = False
        .FormatConditions.Add Type:=xlExpression, Formula1:="=K2>0"
        .FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 5296274
            .TintAndShade = 0
        End With
    
        .FormatConditions(1).StopIfTrue = False
    
    End With
    
    'Apply percent formatting for the max-min analysis
    With Range(Cells(2, 17), Cells(3, 17))
    
        .NumberFormat = "0.00%"
    
    End With

End Sub

Sub formatTotalStockVolume()

    'Apply number formatting to all cells containing data in column L
    Cells(2, 12).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "#,##0"
    
    'Apply number formatting for the max-min analysis
    Cells(4, 17).NumberFormat = "#,##0"

End Sub


