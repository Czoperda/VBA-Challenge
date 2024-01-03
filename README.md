# VBA-Challenge
Repository for Data Science VBA Challenge

Copy of Code:
Sub AddStringToSheets()
    Dim ws As Worksheet
    Dim cell As Range
    Dim strings() As Variant
    Dim i As Integer
    
    'Define cell strings
    strings = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    
    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Set cell strings
        For i = 1 To 4
            Set cell = ws.Cells(1, i + 8)
     
            cell.Value = strings(i - 1)
            
        Next i
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
    Next ws
End Sub

Sub StockInfo()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim Ticker As String
    Dim openprice As Double
    Dim closeprice As Double
    Dim totalVolume As Double
    Dim summaryrow As Long
    
    'Compute for each sheet
    For Each ws In ThisWorkbook.Sheets
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    'Intialize variables
    Ticker = ""
    totalVolume = 0
    summaryrow = 2
    
    'Loop through every row
    For i = 2 To lastRow
    
    If ws.Cells(i, 1).Value <> Ticker Then
                'If it's a new ticker, display summary in columns I, J, K, L
                If Ticker <> "" Then
                    ws.Cells(summaryrow, 9).Value = Ticker
                    ws.Cells(summaryrow, 10).Value = yearlychange
                        If yearlychange < 0 Then
                            ws.Cells(summaryrow, 10).Interior.ColorIndex = 3
                        ElseIf yearlychange >= 0 Then
                            ws.Cells(summaryrow, 10).Interior.ColorIndex = 4
                        End If
                    ws.Cells(summaryrow, 11).Value = IIf(openprice <> 0, percentchange, 0)
                          If percentchange < 0 Then
                            ws.Cells(summaryrow, 11).Interior.ColorIndex = 3
                        ElseIf percentchange >= 0 Then
                            ws.Cells(summaryrow, 11).Interior.ColorIndex = 4
                        End If
                    ws.Cells(summaryrow, 12).Value = totalVolume
                    summaryrow = summaryrow + 1
                End If
                
                'Reset variables for the new ticker
                Ticker = ws.Cells(i, 1).Value
                openprice = ws.Cells(i, 3).Value
                closeprice = ws.Cells(i, 6).Value
                totalVolume = 0
                yearlychange = 0
                percentchange = 0
                End If
            
                'Add data for the current ticker
                openprice = ws.Cells(i, 3).Value
                closeprice = ws.Cells(i, 6).Value
                yearlychange = closeprice - openprice
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                percentchange = IIf(openprice <> 0, yearlychange / openprice, 0)
        Next i

    'Display summary for the last ticker row
        If Ticker <> "" Then
            ws.Cells(summaryrow, 9).Value = Ticker
            ws.Cells(summaryrow, 10).Value = yearlychange
                If yearlychange < 0 Then
                            ws.Cells(summaryrow, 10).Interior.ColorIndex = 3
                        ElseIf yearlychange >= 0 Then
                            ws.Cells(summaryrow, 10).Interior.ColorIndex = 4
                        End If
            ws.Cells(summaryrow, 11).Value = IIf(openprice <> 0, percentchange, 0)
                 If percentchange < 0 Then
                            ws.Cells(summaryrow, 11).Interior.ColorIndex = 3
                        ElseIf percentchange >= 0 Then
                            ws.Cells(summaryrow, 11).Interior.ColorIndex = 4
                        End If
            ws.Cells(summaryrow, 12).Value = totalVolume
        End If
    Next ws
End Sub

Sub YearSummary():
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim percentchange As Double
    Dim Ticker As String
    Dim maxpercentchange As Double
    Dim maxpercentTicker As String
    Dim minpercentchange As Double
    Dim minpercentTicker As String
    Dim totalVolume As Double
    Dim maxtotalVolume As Double
    Dim maxtotalVolumeTicker As String
    
    maxpercentagechange = 0
    minpercentchange = 0
    maxtotalVolume = 0
    
    'Compute for each sheet
    For Each ws In ThisWorkbook.Sheets
    
    ' Find the last row with data in column I
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    'Loop through every row
    For i = 2 To lastRow
        Ticker = ws.Cells(i, 9).Value
         ws.Cells(i, 11).NumberFormat = "0.00%"
        percentchange = ws.Cells(i, 11).Value
        totalVolume = ws.Cells(i, 12).Value
        
        'Calculate Greatest % Increase
        If percentchange > maxpercentchange Then
            maxpercentchange = percentchange
            maxpercentTicker = Ticker
        End If
        
        'Calculate Greatest % Decrease
        If percentchange < minpercentchange Then
            minpercentchange = percentchange
            minpercentTicker = Ticker
        End If
        
        'Calculate Total Volume
        If totalVolume > maxtotalVolume Then
            maxtotalVolume = totalVolume
            maxtotalVolumeTicker = Ticker
        End If
    
     Next i
     
    'Display results
    ws.Cells(2, 16).Value = maxpercentTicker
    ws.Cells(2, 17).Value = maxpercentchange
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = minpercentTicker
    ws.Cells(3, 17).Value = minpercentchange
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = maxtotalVolumeTicker
    ws.Cells(4, 17).Value = maxtotalVolume
        
Next ws
        

End Sub
    

