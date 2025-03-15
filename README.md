# EXCEL-MACRO-PROGRAM-TO-PERFORM-CASH-FLOW-ESTIMATION
Sub EstimateCashFlow()
Dim ws As Worksheet
Dim lastRow As Long
Dim lastColumn As Long
Dim startColumn As Long
Dim totalInflows As Double
Dim totalOutflows As Double
Dim netCashFlow As Double
Dim col As Integer
' Set the worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")
' Find the last row in column A
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
' Find the last column in row 1
lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
' Start column (assuming first column contains descriptions or dates)
