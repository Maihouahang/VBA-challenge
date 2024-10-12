Attribute VB_Name = "Module1"
Sub analyzestocks()
    
'Instructions
    'Create a script that loops through all the stocks for each quarter and outputs the following information:
        'The ticker symbol
        'Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
        'The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
        'The total stock volume of the stock.
        'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
        'Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every quarter) at once.
    'Note - Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
    
    'Looping through all the worksheets
    Dim ws As Worksheet
        For Each ws In Worksheets
    
    'Labeling
    Dim ticker As String
    Dim totalstockvolume As Double
    Dim openprice As Double
    Dim closeprice As Double
    Dim quarterlychange As Double
    Dim percentchange As Double
    
    Dim summaryrow As Long
    Dim lastRow As Long
    Dim quartelychangerow As Long
    
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatesttotalvolume As Double
  
    'Naming headers
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Quartely Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    'Setting initial values
    totalstockvolume = 0
    summaryrow = 2
    
    openprice = ws.Cells(2, 3).Value

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'Looping all the data
        For i = 2 To lastRow
            ' If ticker next ticker is different from current ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker
                ticker = ws.Cells(i, 1).Value
 
                ' calculating quartely change
                closeprice = ws.Cells(i, 6).Value

                quartelychange = closeprice - openprice
 
                ' Calculating percent change
                percentchange = quartelychange / openprice
                
                ' adding total stock volume
                totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
            
            ws.Range("I" & summaryrow).Value = ticker
            ws.Range("J" & summaryrow).Value = quartelychange
            ws.Range("K" & summaryrow).Value = percentchange
            ws.Range("K" & summaryrow).NumberFormat = "0.00%"
            ws.Range("L" & summaryrow).Value = totalstockvolume

        'resetting for next ticker
        totalstockvolume = 0
        summaryrow = summaryrow + 1
        openprice = ws.Cells(i + 1, 3)

        'If ticker of the next cell is the same as the current cell
        Else
            totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value

        End If
        
    Next i
        
        quartelychangerow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
    'Color coding quarterly change
    For j = 2 To quartelychangerow
        'If greater than or equal to zero then green
        If ws.Cells(j, 10) >= 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
        'If less than 0 then red
        ElseIf ws.Cells(j, 10) < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
            
        End If
    
    Next j
             
        'Finding the greatest increase and decrease % and the greatest total volume
        For j = 2 To quartelychangerow
            
            'greatest increase
            If ws.Cells(j, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & quartelychangerow)) Then
                ws.Range("Q2").Value = ws.Cells(j, 11).Value
                ws.Range("Q2").NumberFormat = "0.00%"
                ws.Range("P2").Value = ws.Cells(j, 9).Value
                
            'greatest decrease
            ElseIf ws.Cells(j, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & quartelychangerow)) Then
                ws.Range("Q3").Value = ws.Cells(j, 11).Value
                ws.Range("Q3").NumberFormat = "0.00%"
                ws.Range("P3").Value = ws.Cells(j, 9).Value
                
            'greatest total volume
            ElseIf ws.Cells(j, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & quartelychangerow)) Then
                ws.Range("Q4").Value = ws.Cells(j, 12).Value
                ws.Range("P4").Value = ws.Cells(j, 9).Value
                
            End If
            
        Next j
   
    Next ws

End Sub
