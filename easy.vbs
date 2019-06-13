Option Explicit

Sub vbahomeworkeasy()

Dim sheet As Worksheet
Dim ticker As String
Dim totalstock As String
Dim i As Double
Dim volume As Double
Dim tickername As String
Dim tickersummarytable As Integer
Dim lastrow As Double

ticker = "Ticker"
totalstock = "Total Stock Volume"
volume = 0


    For Each sheet In ActiveWorkbook.Worksheets
        sheet.Range("I1").Cells = ticker
        sheet.Range("J1").Cells = totalstock
        
        lastrow = sheet.Cells(Rows.Count, 1).End(xlUp).Row
        tickersummarytable = 2
        
        For i = 2 To lastrow
        
            If sheet.Cells(i, 1).Value <> sheet.Cells(i + 1, 1).Value Then
        
            volume = volume + sheet.Cells(i, 7).Value
            
            tickername = sheet.Cells(i, 1).Value
            
            sheet.Range("I" & tickersummarytable).Value = tickername
            sheet.Range("J" & tickersummarytable).Value = volume
            
            volume = 0
            tickersummarytable = tickersummarytable + 1
            
            Else
            
            volume = volume + sheet.Cells(i, 7).Value
       
            End If
            
        Next i
        
   Next sheet
   


End Sub