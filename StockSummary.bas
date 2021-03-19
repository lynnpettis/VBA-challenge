Attribute VB_Name = "Module1"
Sub SummarizeStockData()

    ' Formulas to be used in the module
    ' Compute change in value: ChangeInValue = FinalValue - InitialValue
    ' Compute percent change in value: PercentChange = (FinalValue - InitialValue) / InitialValue
        
    For Each Worksheet In Worksheets
    
        'Initialize variables, do all even though it is not needed
        Dim TotalStockVolume, PercentChange As Double
        
        Ticker = " "
        InitialValue = 0
        FinalValue = 0
        ChangeInValue = 0
        PercentChange = 0
        TotalStockVolume = 0
        InputRow = 2
        OutputRow = 1
        Done = False
        LastRow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
                
        ' Initialize new column headers
        Worksheet.Cells(OutputRow, 9).Value = "Ticker"
        Worksheet.Cells(OutputRow, 10).Value = "Yearly Change"
        Worksheet.Cells(OutputRow, 11).Value = "Percent Change"
        Worksheet.Cells(OutputRow, 12).Value = "Total Stock Volume"
        
        ' This loop will process the data on each sheet in the workbook
        
        ' Prime Variables with first row of data
        Ticker = Worksheet.Cells(InputRow, 1).Value
        InitialValue = Worksheet.Cells(InputRow, 3).Value
        TotalStockVolume = Worksheet.Cells(InputRow, 7).Value
        
        While InputRow <= LastRow
        
            If Ticker = Worksheet.Cells(InputRow + 1, 1) Then
                ' Capture next row of data
                FinalValue = Worksheet.Cells(InputRow + 1, 6)
                TotalStockVolume = TotalStockVolume + Worksheet.Cells(InputRow + 1, 7).Value
            Else
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
                
                ' Format the color of the Yearly Change column so the interior
                ' is Green for a positive change and Red for a negative change,
                ' and White for a 0 change
                If ChangeInValue > 0 Then
                    Worksheet.Cells(OutputRow, 10).Interior.ColorIndex = 4
                    Worksheet.Cells(OutputRow, 10).Font.ColorIndex = 1
                ElseIf ChangeInValue < 0 Then
                    Worksheet.Cells(OutputRow, 10).Interior.ColorIndex = 3
                    Worksheet.Cells(OutputRow, 10).Font.ColorIndex = 1
                Else
                    Worksheet.Cells(OutputRow, 10).Interior.ColorIndex = 2
                    Worksheet.Cells(OutputRow, 10).Font.ColorIndex = 1
                End If
                
                ' Set the format of the Percent Change column to display the values as percentages
                Worksheet.Cells(OutputRow, 11).NumberFormat = "0.00%"
                                
                ' Capture Inital Data for next Stock
                Ticker = Worksheet.Cells(InputRow + 1, 1).Value
                InitialValue = Worksheet.Cells(InputRow + 1, 3).Value
                TotalStockVolume = Worksheet.Cells(InputRow + 1, 7).Value
            End If
            InputRow = InputRow + 1
        Wend
        
    Next Worksheet
End Sub
