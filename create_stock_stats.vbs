VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub create_stock_stats()

Dim flag As Boolean
Dim rowNum As Long
Dim incRow As Long
Dim decRow As Long
Dim volRow As Long
Dim lastRow As Long
Dim startNum As Long
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
    rowNum = 2
    startNum = 2
    incRow = 2
    decRow = 2
    volRow = 2
    
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox (lastRow)
    
    Dim totalVol As Integer
    For i = 2 To lastRow
        If Cells(startNum, 1) <> Cells(i, 1) Then
            Cells(rowNum, 9) = Cells(startNum, 1)
            Cells(rowNum, 10) = Cells(i - 1, 6) - Cells(startNum, 3)
            Cells(rowNum, 11) = FormatPercent((Cells(rowNum, 10) / Cells(startNum, 3)))
            Cells(rowNum, 12) = Application.WorksheetFunction.Sum(Range("G" & startNum & ":G" & i - 1))
            
            If Cells(rowNum, 10) > 0 Then
                Cells(rowNum, 10).Interior.ColorIndex = 4
                Cells(rowNum, 11).Interior.ColorIndex = 4
            Else
                Cells(rowNum, 10).Interior.ColorIndex = 3
                Cells(rowNum, 11).Interior.ColorIndex = 3
            End If
        
            If Cells(rowNum, 11) > Cells(incRow, 11) Then
                incRow = rowNum
            End If
            
            If Cells(rowNum, 11) < Cells(decRow, 11) Then
                decRow = rowNum
            End If
                
            If Cells(rowNum, 12) > Cells(volRow, 12) Then
                volRow = rowNum
            End If
            
            rowNum = rowNum + 1
            startNum = i
        End If
    Next i
    
    
    Cells(2, 15) = "Greatest % increase"
    Cells(3, 15) = "Greatest % decrease"
    Cells(4, 15) = "Greatest Total Volume"
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
    
    Cells(2, 16) = Cells(incRow, 9)
    Cells(2, 17) = FormatPercent(Cells(incRow, 11))
    
    Cells(3, 16) = Cells(decRow, 9)
    Cells(3, 17) = FormatPercent(Cells(decRow, 11))
    
    Cells(4, 16) = Cells(volRow, 9)
    Cells(4, 17) = Cells(volRow, 12)
    
    Columns("I:Q").AutoFit
Next ws


End Sub

