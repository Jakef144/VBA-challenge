Attribute VB_Name = "Module1"
Sub Challengetwo()
'Place for all variables
Dim ws As Worksheet
Dim ticker As String
Dim lastrow As Long
Dim OutputRow As Long
Dim i As Long
Dim dateCell As Date
Dim openPrice As Double
Dim closePrice As Double
Dim Volume As Double
Dim quarterName As String
Dim quarterlyData As Object
Dim key As Variant
Dim ExistingData As Variant
Dim Startprice As Double
Dim endPrice As Double
Dim percentageChange As Double
Dim totalVolume As Double
Dim maxPercentagechange As Double
Dim minpercentagechange As Double
Dim maxVolume As Double
Dim maxIncreaseTicker  As String
Dim maxDecreaseTicker As String
Dim maxVolumeTicker As String

'Added looping through all ws
For Each ws In ThisWorkbook.Worksheets
'Resetting the data
    Set quarterlyData = CreateObject("Scripting.Dictionary")
    
    'This was added by GPT when addressing 9999999 errors
    maxPercentagChange = -999999 ' Reset for each sheet
    minpercentagechange = 9999999  ' Reset for each sheet
    maxVolume = 0 ' Reset for each sheet
    
lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

'Loop through column A
For i = 2 To lastrow
ticker = ws.Cells(i, 1).Value
dateCell = ws.Cells(i, 2).Value
Volume = ws.Cells(i, 7).Value

'Checking date
If IsDate(dateCell) Then
 openPrice = ws.Cells(i, 3).Value
 closePrice = ws.Cells(i, 6).Value

' Determine the quarter
                quarterName = Year(dateCell) & "-Q" & Application.WorksheetFunction.RoundUp(Month(dateCell) / 3, 0)
                
                ' Create a unique key for the dictionary using ticker and quarter
                If Not quarterlyData.exists(ticker & "_" & quarterName) Then
                    ' Reset for the start and end of the quarter, and initialize volume
                    quarterlyData.Add ticker & "_" & quarterName, Array(openPrice, closePrice, dateCell, dateCell, Volume)
                Else
                    ' Update existing quarter information
                    ExistingData = quarterlyData(ticker & "_" & quarterName)
                    
                    ' Update only if the new start date is earlier or end date is later
                    If dateCell < ExistingData(2) Then
                        ExistingData(0) = openPrice ' Update start price
                        ExistingData(2) = dateCell  ' Update start date
                    End If
                    If dateCell > ExistingData(3) Then
                        ExistingData(1) = closePrice ' Update end price
                        ExistingData(3) = dateCell   ' Update end date
                    End If
                    
                    ' Add the current volume to the total volume
                    ExistingData(4) = ExistingData(4) + Volume
                    
                    ' Save the updated record back to the dictionary
                    quarterlyData(ticker & "_" & quarterName) = ExistingData
                End If
            End If
        Next i
    
        ' Output quarterly data for the current sheet
        OutputRow = 2 ' Start writing tickers and changes from row 2 in each sheet
        
        ' Headers for each sheet
        ws.Cells(1, 9).Value = "Ticker"             ' Header in I1
        ws.Cells(1, 10).Value = "Quarterly Change"  ' Header in J1
        ws.Cells(1, 11).Value = "Percentage Change" ' Header in K1
        ws.Cells(1, 12).Value = "Total Stock Volume" ' Header in L1
        
        ' Output each ticker and corresponding quarterly change separately
        For Each key In quarterlyData.keys
            ' Split the key to separate ticker and quarter for clarity
            ticker = Split(key, "_")(0)
            Startprice = quarterlyData(key)(0)
            endPrice = quarterlyData(key)(1)
            totalVolume = quarterlyData(key)(4)
            
            ' Calculate the percentage change
            If Startprice <> 0 Then
                percentageChange = ((endPrice - Startprice) / Startprice) * 100
            Else
                percentageChange = 0 ' Handle division by zero if start price is zero
            End If
            
            ' Output the ticker, the change, the percentage change, and the total volume
            ws.Cells(OutputRow, 9).Value = ticker
            ws.Cells(OutputRow, 10).Value = endPrice - Startprice
            ws.Cells(OutputRow, 11).Value = Format(percentageChange, "0.00") & "%" ' Format as a percentage
            ws.Cells(OutputRow, 12).Value = totalVolume ' Output the total volume
            
            OutputRow = OutputRow + 1
            
            ' Track the greatest % increase, decrease, and total volume for the current sheet
            If percentageChange > maxPercentagechange Then
                maxPercentagechange = percentageChange
                maxIncreaseTicker = ticker
            End If
            
            If percentageChange < minpercentagechange Then
                minpercentagechange = percentageChange
                maxDecreaseTicker = ticker
            End If
            
            If totalVolume > maxVolume Then
                maxVolume = totalVolume
                maxVolumeTicker = ticker
            End If
        Next key
        
        ' Output the greatest values for the current sheet
        ws.Cells(2, 14).Value = "Greatest % Increase"  ' N2
        ws.Cells(3, 14).Value = "Greatest % Decrease"  ' N3
        ws.Cells(4, 14).Value = "Greatest Total Volume" ' N4
        
        ws.Cells(1, 15).Value = "Ticker"  ' Header in O1
        ws.Cells(1, 16).Value = "Value"   ' Header in P1
        
        ' Fill in the greatest % increase, % decrease, and greatest volume for the current sheet
        ws.Cells(2, 15).Value = maxIncreaseTicker
        ws.Cells(2, 16).Value = Format(maxPercentagechange, "0.00") & "%"
        
        ws.Cells(3, 15).Value = maxDecreaseTicker
        ws.Cells(3, 16).Value = Format(minpercentagechange, "0.00") & "%"
        
        ws.Cells(4, 15).Value = maxVolumeTicker
        ws.Cells(4, 16).Value = maxVolume

        ' Apply Conditional Formatting to Quarterly Change and Percentage Change Columns
        ' Apply formatting to the Quarterly Change column (Column J)
        With ws.Range("J2:J" & OutputRow - 1)
            ' Clear existing conditional formatting
            .FormatConditions.Delete
            
            ' Add red formatting for values less than or equal to 0
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, Formula1:="=0"
            .FormatConditions(1).Interior.Color = RGB(255, 0, 0) ' Red background
            
            ' Add green formatting for values greater than 0
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(2).Interior.Color = RGB(0, 255, 0) ' Green background
        End With
        
        ' Apply formatting to the Percentage Change column (Column K)
        With ws.Range("K2:K" & OutputRow - 1)
            ' Clear existing conditional formatting
            .FormatConditions.Delete
            
            ' Add red formatting for values less than or equal to 0
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, Formula1:="=0"
            .FormatConditions(1).Interior.Color = RGB(255, 0, 0) ' Red background
            
            ' Add green formatting for values greater than 0
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(2).Interior.Color = RGB(0, 255, 0) ' Green background
        End With
        
    Next ws
    
    MsgBox "Quarterly and yearly data have been processed and summarized.", vbInformation
End Sub
