Attribute VB_Name = "Module1"
Sub ticker()

Dim i As Double
Dim summarycount As Double
Dim lastrow As Double
Dim ticker As String
Dim opens As Double
Dim closes As Double
Dim nextticker As String
Dim prevticker As String
Dim volumes As Double

lastrow = Range("A1").End(xlDown).Row

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % increase"
Cells(3, 15).Value = "Greatest % decrease"
Cells(4, 15).Value = "Greatest total volume"

'ticker = 1
'dates = 2
'opens = 3
'closes = 6
'volumes = 7

summarycount = 2

For i = 2 To lastrow
    
    prevticker = Cells(i - 1, 1).Value
    ticker = Cells(i, 1).Value
    nextticker = Cells(i + 1, 1).Value
    
    volumes = volumes + Cells(i, 7).Value
    
    
    'finds first row
    If prevticker <> ticker Then
    opens = Cells(i, 3).Value
    
    
    
    'finds last row
    ElseIf ticker <> nextticker Then
    closes = Cells(i, 6).Value
    
    
    
    Cells(summarycount, 9).Value = ticker
    Cells(summarycount, 12).Value = volumes
    Cells(summarycount, 10).Value = closes - opens
    Cells(summarycount, 11).Value = (closes - opens) / opens * 100
    
    summarycount = summarycount + 1
    
    volumes = 0
    
    End If
    
        


Next i




End Sub

