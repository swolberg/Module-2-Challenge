Attribute VB_Name = "Module3"
Sub CalculateGreatestChanges()
    Dim ws As Worksheet
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim startRow As Long, endRow As Long
    Dim lastRow As Long
    Dim outputRow As Long
    Dim quarterlyChange As Double
    Dim percentageChange As Double

    ' Variables to track greatest increase, decrease, and volume
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    ' Loop through each worksheet (assuming quarterly sheets)
    For Each ws In Worksheets
        If ws.Name <> "Output" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Find the last row with data
            
            ' Initialize the tracking variables for each sheet
            greatestIncrease = -9999999
            greatestDecrease = 9999999
            greatestVolume = 0
            
            ' Initialize the first row for the ticker and quarter
            startRow = 2
            ticker = ws.Cells(startRow, 1).Value
            
            ' Add headers if they don't exist already in Columns I to L
            If ws.Cells(1, 9).Value = "" Then
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 10).Value = "Quarterly Change"
                ws.Cells(1, 11).Value = "Percentage Change"
                ws.Cells(1, 12).Value = "Total Volume"
            End If
            
            outputRow = 2 ' Start adding the output after the last stock data row
            
            Do While startRow <= lastRow
                ' Find the range for each ticker
                endRow = startRow
                Do While ws.Cells(endRow, 1).Value = ticker And endRow <= lastRow
                    endRow = endRow + 1
                Loop
                endRow = endRow - 1 ' Adjust to the last row for the current ticker
                
                ' Calculate values for the quarter
                openPrice = ws.Cells(startRow, 3).Value
                closePrice = ws.Cells(endRow, 6).Value
                quarterlyChange = closePrice - openPrice
                percentageChange = (quarterlyChange / openPrice) * 100
                
                ' Sum up the volume for the entire quarter
                volume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7)))
                
                ' Output the results in the same sheet starting from Column I
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = quarterlyChange
                ws.Cells(outputRow, 11).Value = percentageChange
                ws.Cells(outputRow, 12).Value = volume
                
                ' Conditional formatting for quarterly change
                If quarterlyChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.Color = RGB(144, 238, 144) ' Light green for positive
                ElseIf quarterlyChange < 0 Then
                    ws.Cells(outputRow, 10).Interior.Color = RGB(255, 182, 193) ' Light red for negative
                End If
                
                ' Conditional formatting for percentage change
                If percentageChange > 0 Then
                    ws.Cells(outputRow, 11).Interior.Color = RGB(144, 238, 144) ' Light green for positive
                ElseIf percentageChange < 0 Then
                    ws.Cells(outputRow, 11).Interior.Color = RGB(255, 182, 193) ' Light red for negative
                End If
                
                ' Check for greatest % increase
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    greatestIncreaseTicker = ticker
                End If
                
                ' Check for greatest % decrease
                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    greatestDecreaseTicker = ticker
                End If
                
                ' Check for greatest total volume
                If volume > greatestVolume Then
                    greatestVolume = volume
                    greatestVolumeTicker = ticker
                End If
                
                ' Move to the next row and reset for the next ticker
                outputRow = outputRow + 1
                startRow = endRow + 1
                If startRow <= lastRow Then
                    ticker = ws.Cells(startRow, 1).Value
                End If
            Loop
            
            ' Output the greatest increase, decrease, and total volume starting from O2
            ws.Cells(2, 15).Value = "Criteria"
            ws.Cells(2, 16).Value = "Ticker"
            ws.Cells(2, 17).Value = "Value"

            ws.Cells(3, 15).Value = "Greatest % Increase"
            ws.Cells(3, 16).Value = greatestIncreaseTicker
            ws.Cells(3, 17).Value = greatestIncrease

            ws.Cells(4, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 16).Value = greatestDecreaseTicker
            ws.Cells(4, 17).Value = greatestDecrease

            ws.Cells(5, 15).Value = "Greatest Total Volume"
            ws.Cells(5, 16).Value = greatestVolumeTicker
            ws.Cells(5, 17).Value = greatestVolume
            
            ' Conditional formatting for summary: Greatest Increase/Decrease
            If greatestIncrease > 0 Then
                ws.Cells(3, 17).Interior.Color = RGB(144, 238, 144) ' Light green for positive
            ElseIf greatestIncrease < 0 Then
                ws.Cells(3, 17).Interior.Color = RGB(255, 182, 193) ' Light red for negative
            End If
            
            If greatestDecrease > 0 Then
                ws.Cells(4, 17).Interior.Color = RGB(144, 238, 144) ' Light green for positive
            ElseIf greatestDecrease < 0 Then
                ws.Cells(4, 17).Interior.Color = RGB(255, 182, 193) ' Light red for negative
            End If
        End If
    Next ws
    
    MsgBox "Greatest changes calculated and added to each sheet!"
End Sub

