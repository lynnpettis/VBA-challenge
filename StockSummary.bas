Attribute VB_Name = "Module1"
Sub SummarizeStockData()

    ' Formulas to be used in the module
    ' Compute change in value: ChangeInValue = FinalValue - InitialValue
    ' Compute percent change in value: PercentChange = (FinalValue - InitialValue) / InitialValue
    '   This can be reduced to ChangeInValue / InitialValue
        
    For Each Worksheet In Worksheets
    
        'Initialize variables, do all even though it is not needed
        Ticker = " "
        InitialValue = 0
        FinalValue = 0
        ChangeInValue = 0
        PercentChange = 0
        TotalStockVolume = 0
        InputRow = 2
        OutputRow = 1
        LastRow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
                
        ' Initialize new column headers
        Worksheet.Cells(OutputRow, 9).Value = "Ticker"
        Worksheet.Cells(OutputRow, 10).Value = "Yearly Change"
        Worksheet.Cells(OutputRow, 11).Value = "Percent Change"
        Worksheet.Cells(OutputRow, 12).Value = "Total Stock Volume"
        
        ' For the Bonus Section
        ' Initialize section to hold Stocks with
        ' the greatest % increase, greatest % decrease,
        ' and greatest total volume on each sheet
        
        Worksheet.Cells(1, 16).Value = "Ticker"
        Worksheet.Cells(1, 17).Value = "Value"
        Worksheet.Cells(2, 15).Value = "Greatest % Increase"
        Worksheet.Cells(3, 15).Value = "Greatest % Decrease"
        Worksheet.Cells(4, 15).Value = "Greatest Total Volume"
        Worksheet.Cells(2, 17).Value = 0
        Worksheet.Cells(2, 17).NumberFormat = "0.00%"
        Worksheet.Cells(3, 17).Value = 0
        Worksheet.Cells(3, 17).NumberFormat = "0.00%"
        Worksheet.Cells(4, 17).Value = 0
        
        ' This loop will process the data on each sheet in the workbook
        
        ' Prime Variables with first row of data
        ' Capturing all the required data insures that
        ' you have all the data should a given Stock Ticker be
        ' a single row of data
        Ticker = Worksheet.Cells(InputRow, 1).Value
        InitialValue = Worksheet.Cells(InputRow, 3).Value
        FinalValue = Worksheet.Cells(InputRow, 6)
        TotalStockVolume = Worksheet.Cells(InputRow, 7).Value
        
        While InputRow <= LastRow
        
            If Ticker = Worksheet.Cells(InputRow + 1, 1) Then
                ' Capture next requesit data from the next row of data
                FinalValue = Worksheet.Cells(InputRow + 1, 6)
                TotalStockVolume = TotalStockVolume + Worksheet.Cells(InputRow + 1, 7).Value
            Else
                ' The value of Ticker is changing in the next row,
                ' Calculate Values for current Stock
                ChangeInValue = FinalValue - InitialValue
                
                If InitialValue = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = ChangeInValue / InitialValue
                End If
                ' Output data to spreadsheet
                OutputRow = OutputRow + 1
                Worksheet.Cells(OutputRow, 9).Value = Ticker
                Worksheet.Cells(OutputRow, 10).Value = ChangeInValue
                Worksheet.Cells(OutputRow, 11).Value = PercentChange
                Worksheet.Cells(OutputRow, 12).Value = TotalStockVolume
                
                ' This where we will capture the data for the Bonus
                ' section.  This code will allow the maximum or minimum
                ' values and the ticker value to "bubble up" to the top,
                ' allowing for the capture of the first instance of the
                ' Stocks with the greatest % increase, greatest % decrease,
                ' and the greatest total volume.
                
                ' Greatest % Increase
                If PercentChange > Worksheet.Cells(2, 17).Value Then
                    Worksheet.Cells(2, 16).Value = Ticker
                    Worksheet.Cells(2, 17).Value = PercentChange
                End If
                
                ' Greatest % Decrease
                If PercentChange < Worksheet.Cells(3, 17).Value Then
                    Worksheet.Cells(3, 16).Value = Ticker
                    Worksheet.Cells(3, 17).Value = PercentChange
                End If
                
                ' Greatest Total Volume
                If TotalStockVolume > Worksheet.Cells(4, 17).Value Then
                    Worksheet.Cells(4, 16).Value = Ticker
                    Worksheet.Cells(4, 17).Value = TotalStockVolume
                End If
               
                ' Format the color of the Yearly Change column so the interior
                ' is Green for a positive change and Red for a negative change,
                ' and White for a 0 change
                If ChangeInValue > 0 Then
                    Worksheet.Cells(OutputRow, 10).Interior.ColorIndex = 4
                ElseIf ChangeInValue < 0 Then
                    Worksheet.Cells(OutputRow, 10).Interior.ColorIndex = 3
                Else
                    Worksheet.Cells(OutputRow, 10).Interior.ColorIndex = 2
                End If
                
                ' Ensure the font is set to Black
                Worksheet.Cells(OutputRow, 10).Font.ColorIndex = 1
                
                ' Set the format of the Percent Change column to display the values as percentages
                Worksheet.Cells(OutputRow, 11).NumberFormat = "0.00%"
                                
                ' Capture Inital Data for next Stock
                Ticker = Worksheet.Cells(InputRow + 1, 1).Value
                InitialValue = Worksheet.Cells(InputRow + 1, 3).Value
                FinalValue = Worksheet.Cells(InputRow + 1, 6).Value
                TotalStockVolume = Worksheet.Cells(InputRow + 1, 7).Value
            End If
            
            ' Increment the InputRow
            InputRow = InputRow + 1
        Wend
        
        ' Autofit data in new columns
        Worksheet.Columns("O:Q").AutoFit
        Worksheet.Columns("I:L").AutoFit
        
    Next Worksheet
End Sub
