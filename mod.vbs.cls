VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub year_change()
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
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Dim Summary_Table_Row As Long
  Summary_Table_Row = 2

For i = 2 To lastrow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

ticker = Cells(i, 1).Value
volume = volume + Cells(i, 7).Value

Range("K" & Summary_Table_Row).Value = ticker
Range("N" & Summary_Table_Row).Value = volume


    If Cells(i, 2).Value = first_day Then
    open_day = Cells(i, 3).Value
    ElseIf Cells(i, 2).Value = last_day Then
    close_day = Cells(i, 6).Value
    
    End If
    
yearly_change = close_day - open_day
Range("L" & Summary_Table_Row).Value = yearly_change
    
    If Range("L" & Summary_Table_Row).Value > 0 Then
    Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
    Else
    Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
    
    End If
    


    If open_day = 0 Then
    percent_change = " "
    Else
    percent_change = yearly_change / open_day
    
    End If

Range("M" & Summary_Table_Row).Value = FormatPercent(percent_change, 1)

    
Summary_Table_Row = Summary_Table_Row + 1
volume = 0

Else
volume = volume + Cells(i, 7).Value
If Cells(i, 2).Value = first_day Then
    open_day = Cells(i, 3).Value
    ElseIf Cells(i, 2).Value = last_day Then
    close_day = Cells(i, 6).Value
    
    End If
    
End If

Next i




End Sub




