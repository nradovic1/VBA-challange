Attribute VB_Name = "Module1"
Sub year_change()

Dim ws As Worksheet

For Each ws In Worksheets

ws.Range("K1,R1").Value = "Ticker"
ws.Range("L1").Value = "Yearly Change"
ws.Range("M1").Value = "Percent Change"
ws.Range("N1").Value = "Total Stock Volume"

Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double

Dim volume As LongLong
volume = 0

Dim first_day As Double
Dim last_day As Double

Dim open_day As Double
Dim close_day As Double


first_day = WorksheetFunction.Min(ws.Range("B:B"))
last_day = WorksheetFunction.Max(ws.Range("B:B"))

yearly_change = close_day - open_day

Dim lastrow As LongLong
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim Summary_Table_Row As Long
  Summary_Table_Row = 2

For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

ticker = ws.Cells(i, 1).Value
volume = volume + ws.Cells(i, 7).Value

ws.Range("K" & Summary_Table_Row).Value = ticker
ws.Range("N" & Summary_Table_Row).Value = volume


    If ws.Cells(i, 2).Value = first_day Then
    open_day = ws.Cells(i, 3).Value
    ElseIf ws.Cells(i, 2).Value = last_day Then
    close_day = ws.Cells(i, 6).Value
    
    End If
    
yearly_change = close_day - open_day
ws.Range("L" & Summary_Table_Row).Value = yearly_change
    
    If ws.Range("L" & Summary_Table_Row).Value > 0 Then
    ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
    Else
    ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
    
    End If
    


    If open_day = 0 Then
    percent_change = " "
    Else
    percent_change = yearly_change / open_day
    
    End If

ws.Range("M" & Summary_Table_Row).Value = FormatPercent(percent_change, 1)

    
Summary_Table_Row = Summary_Table_Row + 1
volume = 0

Else
volume = volume + ws.Cells(i, 7).Value
If ws.Cells(i, 2).Value = first_day Then
    open_day = ws.Cells(i, 3).Value
    ElseIf ws.Cells(i, 2).Value = last_day Then
    close_day = ws.Cells(i, 6).Value
    
    End If
    
End If

Next i

ws.Range("Q2").Value = "Greatest % Increase"
ws.Range("Q3").Value = "Greatest % Decrease"
ws.Range("Q4").Value = "Greatest Total Volume"
ws.Range("S1").Value = "Value"

Dim greatest As Double
Dim lowest As Double
Dim best_volume As LongLong


greatest = WorksheetFunction.Max(ws.Range("M:M"))
lowest = WorksheetFunction.Min(ws.Range("M:M"))
best_volume = WorksheetFunction.Max(ws.Range("N:N"))

For i = 2 To lastrow

    If ws.Cells(i, 13).Value = greatest Then
    ws.Range("R2").Value = ws.Cells(i, 11).Value
    ws.Range("S2").Value = FormatPercent(ws.Cells(i, 13).Value, 1)
    
    ElseIf ws.Cells(i, 13).Value = lowest Then
    ws.Range("R3").Value = ws.Cells(i, 11).Value
    ws.Range("S3").Value = FormatPercent(ws.Cells(i, 13).Value, 1)
    
    ElseIf ws.Cells(i, 14).Value = best_volume Then
    ws.Range("R4").Value = ws.Cells(i, 11).Value
    ws.Range("S4").Value = ws.Cells(i, 14).Value
    
    
    End If
          
 Next i





Next ws


End Sub







