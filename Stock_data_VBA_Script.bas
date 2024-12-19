Attribute VB_Name = "Module1"
Sub AnalysisOfGeneratedStockMarketData()

' The requirement is to make the VBA code in this module to run through all the sheets in the provided workbook.
' As suggested in the instructions, the VBA scripting was done the smaller dataset and that allowed a faster test.
' The first step is to name the variables.

    Dim ws As Worksheet
    Dim Ticker As String
    Dim Quarterly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As LongLong
    Dim Quarterly_Open As Double
    Dim Quarterly_Close As Double
    Dim Quarterly_Volume As LongLong
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Total_Volume As LongLong
    Dim Greatest_Increase_Ticker As String
    Dim Greatest_Decrease_Ticker As String
    Dim Greatest_Total_Volume_Ticker As String
    Dim LastRow As Long
    Dim Summary_Table_Row As Long
    
    
' The following line helped apply the variables to each worksheet.
    
    For Each ws In ThisWorkbook.Worksheets
    
' Creating the two new tables (rows and columns) using ranges to display the required results.
' The first new table is the table for the output of the loops. The requied columns were ticker, quarterly change, percent change, and total stock volume.
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
' Auto-format the columns for the new output/summary table to fit with a resized coulumn.
' This method is modified from the Range.Autofit method on microsoft's website. Visit: https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit

    ws.Range("I1").EntireColumn.AutoFit
    ws.Range("J1").EntireColumn.AutoFit
    ws.Range("K1").EntireColumn.AutoFit
    ws.Range("L1").EntireColumn.AutoFit
    
' The second table is used to display the Greatest % Increase, Greatest % Decrease and the Greatest Total Volume.

    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
   
' Auto-format the columns for the above table.
' Same formatting as above was applied.

    ws.Range("O1").EntireColumn.AutoFit
    
' Setting initial/default values to variables.
    
    Total_Stock_Volume = 0
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Total_Volume = 0
    Summary_Table_Row = 2
    
' Setting the last row.
' The following line is discovered in the solved star_counter_with_VBA_solution workbook provided for the VBA part of the course.

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
' Looping

    For i = 2 To LastRow
    
' Set Quarely Open for the first entry of each ticker

    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        Quarterly_Open = ws.Cells(i, 3).Value
        
    End If
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ' Ticker
            
            Ticker = ws.Cells(i, 1).Value
            
            ' Total Stock Volume
            
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
            ' Quarterly Close
            
            Quarterly_Close = ws.Cells(i, 6).Value
            
            ' Populate ticker name
            
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            
            ' Qaurterly Change
            
            Quarterly_Change = Quarterly_Close - Quarterly_Open
            
            ' Populate Quarterly Change as a double.
            
            ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
            
             
            ' Conditional Formatting based on the quarterly changes usig the RGB Function as shown in the microsoft website.
            ' Visit: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/rgb-function
            ' The required task was to use conditional formatting that will highlight positive change in green and negative change in red.
            ' Although no explicitly mentioned in the instructions, the provided picture for the challenge showed that when the value = 0, the interior of the cell is to be left white.
           
            
            If Quarterly_Change > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
                
            ElseIf Quarterly_Change < O Then
            
                ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
                
            Else
            
                ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 255, 255)
            
            End If
            
            ' Percent Change calculation
            
            If Quarterly_Open <> 0 Then
            
                Percent_Change = Quarterly_Change / Quarterly_Open
                
            Else: Percent_Change = 0
            
            End If
                
            ' Populate Percent Change to the designated column
            ' The provided image for the challenge showed that the values of the percent change follow the format "0.00%"
            ' To match the format, the Range.NumberFormat property is used from the microsoft website with minor tweaks to match the need for this dataset.
            ' Visit: https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformat
            
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
            
            ' Total Stock Volume value print
            
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            
            ' Check for greatest increase, decrease and volume
            ' Increase
            
            If Percent_Change > Greatest_Increase Then
                Greatest_Increase = Percent_Change
                Greatest_Increase_Ticker = Ticker
                
            End If
            
            ' Decrease
            
            If Percent_Change < Greatest_Decrease Then
                Greatest_Decrease = Percent_Change
                Greatest_Decrease_Ticker = Ticker
            
            End If
            
            ' Volume
            
            If Total_Stock_Volume > Greatest_Total_Volume Then
                Greatest_Total_Volume = Total_Stock_Volume
                Greatest_Total_Volume_Ticker = Ticker
                
            End If
            
            ' Print to table
            
            ws.Cells(2, 16).Value = Greatest_Increase_Ticker
            ws.Cells(2, 17).Value = Greatest_Increase
            
            ws.Cells(3, 16).Value = Greatest_Decrease_Ticker
            ws.Cells(3, 17).Value = Greatest_Decrease
            
            ws.Cells(4, 16).Value = Greatest_Total_Volume_Ticker
            ws.Cells(4, 17).Value = Greatest_Total_Volume
            
            ' Format the table
            
            ws.Cells(2, 17).NumberFormat = "00.00%"
            ws.Cells(3, 17).NumberFormat = "00.00%"
            
            ' Move to the next row in the summary table
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset variables for the next ticker
            
            Total_Stock_Volume = 0
            Quarterly_Open = ws.Cells(i + 1, 3).Value
            
            Else
            ' Add to tatal stock volume for current ticker
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
            End If
            
        Next i
 
    Next ws

End Sub

