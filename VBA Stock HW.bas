Attribute VB_Name = "Module1"
Sub stock_data()

For Each ws In Worksheets

'---------------------------------------------------------------------------------------
'setup
    
    ws.Range("H1") = "Ticker"
    ws.Range("I1") = "Yearly Change"
    ws.Range("J1") = "Percent Change"
    ws.Range("K1") = "Total Volume"
    
'---------------------------------------------------------------------------------------
'variable setup
    
    'set variables
    Dim ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim total_volume As Double

    'need reference variable to cycle through reference table starting at 2nd row
    Dim x As Integer
    x = 2

    'find last row
    Dim last_row As Double
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

'---------------------------------------------------------------------------------------

    For i = 2 To last_row
        
        'if next ticker is = current ticker, has to add volume to previous stored volume for set
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            'if previous ticker is <> current ticker, will need to grab year open value
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                'set year open variable
                year_open = ws.Cells(i, 3).Value
            
            End If
        
        'if next ticker is <> current ticker, then need to grab variables from that row
        Else
            'set ticker variable
            ticker = ws.Cells(i, 1).Value
            'set year close variable
            year_close = ws.Cells(i, 6).Value
            'set ticker cell variable
            ws.Cells(x, 8).Value = ticker
            'add cell volume to total volume
            total_volume = total_volume + ws.Cells(i, 7).Value
            'set total volume cell variable
            ws.Cells(x, 11).Value = total_volume
            'calculate yearly change
            ws.Cells(x, 9).Value = year_close - year_open
'---------------------------------------------------------------------------------------
'cell formatting
                    If ws.Cells(x, 9).Value > 0 Then
                        ws.Cells(x, 9).Interior.ColorIndex = 4
                    Else
                        ws.Cells(x, 9).Interior.ColorIndex = 3
                    End If
                
'---------------------------------------------------------------------------------------
            
            'avoid /0
            If year_open <> 0 Then
                'Percent change calculation
                ws.Cells(x, 10).Value = ws.Cells(x, 9).Value / year_open
            
            End If
            
            'set % formatting
            ws.Cells(x, 10).Style = "percent"
            'reset total volume
            total_volume = 0
            'cycle to next field in reference table
            x = x + 1
            
        End If
        
    Next i

'---------------------------------------------------------------------------------------
'*challenge*
'---------------------------------------------------------------------------------------
'challenge setup

    ws.Range("N2") = "Greatest Total Volume"
    ws.Range("N3") = "Greatest % Increase"
    ws.Range("N4") = "Greatest % Decrease"
    ws.Range("O1") = "Ticker"
    ws.Range("P1") = "Value"
    
'---------------------------------------------------------------------------------------
'variable setup
    
    Dim greatest_total_volume As Double

    'set gtv variable = 0
    greatest_total_volume = 0


'---------------------------------------------------------------------------------------
'greatest volume loop
    
    For i = 2 To x
    
        If ws.Cells(i, 11).Value > greatest_total_volume Then
            greatest_total_volume = ws.Cells(i, 11).Value
            'add ticker
            ws.Cells(2, 15).Value = ws.Cells(i, 8).Value
            
        End If
        
    ws.Cells(2, 16).Value = greatest_total_volume
        
    Next i

    
'---------------------------------------------------------------------------------------
'variable setup

    Dim percent_increase As Double
    Dim percent_decrease As Double

    'set % variables = 0
    percent_increase = 0
    percent_decrease = 0

'---------------------------------------------------------------------------------------
'greatest % increase/decrease loop
    
    For i = 2 To x
    
        'loop through % change to find greatest increase or decrease
        If ws.Cells(i, 10).Value > percent_increase Then
            percent_increase = ws.Cells(i, 10).Value
            'add ticker
            ws.Cells(3, 15).Value = ws.Cells(i, 8).Value
            
        ElseIf ws.Cells(i, 10).Value < percent_decrease Then
            percent_decrease = ws.Cells(i, 10).Value
            'add ticker
            ws.Cells(4, 15).Value = ws.Cells(i, 8).Value
            
        End If
        
    'enter variable to appropriate cell w/ formatting
    ws.Cells(3, 16).Value = percent_increase
    ws.Cells(4, 16).Value = percent_decrease
    
    Next i
    
'---------------------------------------------------------------------------------------
'% formatting

    ws.Cells(3, 16).Style = "percent"
    ws.Cells(4, 16).Style = "percent"


Next ws

End Sub
