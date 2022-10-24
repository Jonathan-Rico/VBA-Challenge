Attribute VB_Name = "Module2"
Sub tickerdata()
'Variable Declaration'
Dim x As Integer
Dim ws As Worksheet
Dim Ticker_Name As String
Dim Ticker_Total, gtotal As LongLong
Dim Table_Row As Integer
Dim Lastrow As Long
Dim vopen, vclose As Double
Dim gpercent, lpercent As Double
Dim cell As Range, k As Double

'Loop to apply this to all worksheets'
For Each ws In ThisWorkbook.Worksheets()

'Variables being assigned values, creating headers for table where output will be displayed'
    Ticker_Total = 0
    Table_Row = 2
    gpercent = 0
    lpercent = 0
    gtotal = 0
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Decrease"
    ws.Range("O3").Value = "Greatest % Increase"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    vopen = ws.Cells(2, 3).Value
    
'For loop to check original ticker data'
    For i = 2 To Lastrow
        
        'If statement that grabs opening ticker value for the first row of data'
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value And i <> 2 Then
        vopen = ws.Cells(i, 3).Value
        
        End If
        'If statement checking if current row's ticker name is the same as the last ticker name, then outputs total stock value, percent change and yearly change'
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Ticker_Name = ws.Cells(i, 1).Value
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
            vclose = ws.Cells(i, 6).Value
            ws.Range("J" & Table_Row).Value = Ticker_Name
            ws.Range("M" & Table_Row).Value = Ticker_Total
            ws.Range("K" & Table_Row).Value = (vclose - vopen)
            ws.Range("L" & Table_Row).Value = (((vclose - vopen) / (vopen)))
            ws.Range("L" & Table_Row).Value = Format(ws.Range("L" & Table_Row).Value, "0.00%")
            
        'If statement that color codes the cells depending on whether percent change was negative or positive'
            If (vclose - vopen) < 0 Then
                ws.Range("K" & Table_Row).Interior.ColorIndex = 3
                
            ElseIf (vclose - vopen) > 0 Then
                ws.Range("K" & Table_Row).Interior.ColorIndex = 4
                
            End If
        'If statements that compare ticker total, percent change to determine which ticker has the greatest decrease or increase in percentage, and greatest total volume'
            If Ticker_Total > gtotal Then
                gtotal = Ticker_Total
                ws.Range("P4").Value = Ticker_Name
                ws.Range("Q4").Value = gtotal
                End If
            If (((vclose - vopen) / (vopen))) > gpercent Then
                gpercent = (((vclose - vopen) / (vopen)))
                ws.Range("P3").Value = Ticker_Name
                ws.Range("Q3").Value = gpercent
                ws.Range("Q3").Value = Format(ws.Range("Q3").Value, "0.00%")
                End If
            If lpercent > (((vclose - vopen) / (vopen))) Then
                lpercent = (((vclose - vopen) / (vopen)))
                ws.Range("P2").Value = Ticker_Name
                ws.Range("Q2").Value = lpercent
                ws.Range("Q2").Value = Format(ws.Range("Q2").Value, "0.00%")
                End If
                
        'Adding to table row counter to ensure output is inputted into separate rows per ticker, resetting the total volume to zero'
            Table_Row = Table_Row + 1
            
            Ticker_Total = 0
        
        'Else statement that adds to the total volume since the ticker name is the same
        Else
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

        End If

    Next i
    
Next

End Sub
