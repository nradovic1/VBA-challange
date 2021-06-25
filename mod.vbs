Sub year_change()

for each ws in worksheets

Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double

Dim volume As LongLong
volume = 0

Dim first_day As Double
Dim last_day As Double

Dim open_day As Double
Dim close_day As Double


first_day = WorksheetFunction.Min(Range("B:B"))
last_day = WorksheetFunction.Max(Range("B:B"))

yearly_change = close_day - open_day

Dim lastrow As LongLong
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim Summary_Table_Row As Long
  Summary_Table_Row = 2

For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

ticker = ws.Cells(i, 1).Value
volume = volume + Cells(i, 7).Value

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
volume = volume + Cells(i, 7).Value
If wsCells(i, 2).Value = first_day Then
    open_day = ws.Cells(i, 3).Value
    ElseIf ws.Cells(i, 2).Value = last_day Then
    close_day = ws.Cells(i, 6).Value
    
    End If
    
End If

Next i

next ws


End Sub




