Attribute VB_Name = "Module1"
Sub Stock_Volume()

Dim ws As Worksheet
Dim last_row, last_col As Long
Dim summary_table_row As Integer
Dim new_ticker As String
Dim total_vol, first_open, last_close As Double

For Each ws In Worksheets
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row                                                            'find the last row in the page
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    summary_table_row = 2
    total_vol = 0
    first_open = ws.Cells(2, 3).Value                                                                                        'get an "open" value for the very first ticker on the current page
    For r = 2 To last_row
        If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
            new_ticker = ws.Cells(r, 1).Value                                                                                'update the "new ticker" variable
            total_vol = total_vol + ws.Cells(r, 7).Value                                                                  'add to "total valume" value from the current row
            last_close = ws.Cells(r, 6).Value                                                                                  'record the "close" value for the current ticker
            ws.Range("J" & summary_table_row).Value = last_close - first_open                            'to record yearly change
            
                If ws.Range("J" & summary_table_row).Value < 0 Then                                             'format the yearly change column
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                End If
                
                If first_open = 0 Then                                                                                             'to calculate percent change
                    ws.Range("K" & summary_table_row).Value = 0
                Else
                    ws.Range("K" & summary_table_row).Value = (last_close - first_open) / first_open
                End If
                
            ws.Range("K" & summary_table_row).NumberFormat = "0.00%"                                   'to apply % format
            
                If ws.Cells(r + 1, 1).Value <> "" Then
                    first_open = ws.Cells(r + 1, 3).Value                                                                   'get an "open" value for the next ticker
                End If
            
            ws.Range("I" & summary_table_row).Value = new_ticker                                             'record a ticker name to the summary table
            ws.Range("L" & summary_table_row).Value = total_vol                                                'record "total volume" amount to the summary table for the current ticket
            summary_table_row = summary_table_row + 1                                                          'set a new "last row" in the summary table
            total_vol = 0                                                                                                               'reset "total volume" variable to 0 for the next ticker
        Else
            total_vol = total_vol + ws.Cells(r, 7).Value                                                                  'add volume for the current ticker if next row is the same ticker
        End If
    Next r
    

    'Bonus part:
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    Dim ticker As String
    Dim Max, Min, Max_vol As Double
    Max = 0
    Min = 0
    Max_vol = 0
    
    For i = 2 To summary_table_row
        If ws.Cells(i, 11) > Max Then
           Max = ws.Cells(i, 11).Value
           ticker = ws.Cells(i, 9).Value
        End If
    Next i
    ws.Range("P2").Value = ticker
    ws.Range("Q2").Value = Max
    ws.Range("Q2").NumberFormat = "0.00%"
    
    For i = 2 To summary_table_row
        If ws.Cells(i, 11) < Min Then
           Min = ws.Cells(i, 11).Value
           ticker = ws.Cells(i, 9).Value
        End If
    Next i
    ws.Range("P3").Value = ticker
    ws.Range("Q3").Value = Min
    ws.Range("Q3").NumberFormat = "0.00%"
    
    For i = 2 To summary_table_row
        If ws.Cells(i, 12) > Max_vol Then
           Max_vol = ws.Cells(i, 12).Value
           ticker = ws.Cells(i, 9).Value
        End If
    Next i
    ws.Range("P4").Value = ticker
    ws.Range("Q4").Value = Max_vol

Next ws

End Sub

