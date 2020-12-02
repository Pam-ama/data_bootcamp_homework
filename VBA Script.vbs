Sub Stocks():

    Dim current_sheet As Worksheet

    For Each current_sheet In Worksheets
        sheet_name = current_sheet.Name

        Worksheets(sheet_name).Activate

        '' to get last row
        Dim last_row As Long
        last_row = current_sheet.Cells(current_sheet.Rows.Count, "A").End(xlUp).Row

        ''define stock volume
        Dim stock_volume As Double
        stock_volume = 0

        '' define summary row, the row where results will be displayed
        Dim summary_row As Long
        summary_row = 2

        ''to show the headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percentage Change"
        Cells(1, 12).Value = "Stock Volume"

        '' define open value and the starting open value
        Dim open_value As Double
        open_value = Cells(2, 3).Value

        For I = 2 To last_row

            If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
            
            stock_volume = stock_volume + Cells(I, 7).Value
 
            Dim close_value As Double
            close_value = Cells(I, 6).Value

            '' close_value - open_value will be the yearly change
            Cells(summary_row, 10).Value = close_value - open_value
 
            ''formatting the colour changes in the Yearly Change Column
            If Cells(summary_row, 10).Value >= 0 Then
            Cells(summary_row, 10).Interior.ColorIndex = 4
            Else
            Cells(summary_row, 10).Interior.ColorIndex = 3
            End If
 
            ''When calculating percentage change, if open value or close value is 0, will print "nill" in that specific cell
            If close_value = 0 Or open_value = 0 Then
            Cells(summary_row, 11).Value = "nill"
            Else
            Cells(summary_row, 11).Value = (close_value / open_value) - 1
            Columns(11).NumberFormat = "0.00%"
            End If
    
            open_value = Cells(I + 1, 3).Value
 
 
            ''Extracting the ticker and stock volume values
            Cells(summary_row, 9).Value = Cells(I, 1).Value
            Cells(summary_row, 12).Value = stock_volume

            summary_row = summary_row + 1
            
            ''Reset stock volume value
            stock_volume = 0

            Else
            stock_volume = stock_volume + Cells(I, 7).Value

            End If
        Next I
    Next


End Sub
