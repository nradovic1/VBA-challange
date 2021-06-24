Attribute VB_Name = "Module1"
Sub stocks()

Dim ws As Worksheet

Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double




Dim volume As LongLong
volume = 0
Dim open_day As Double
open_day = 0
Dim close_day As Double
close_day = 0

yearly_change = close_day - open_day

Dim first_day As Double
Dim last_day As Double


Dim rng As Range
Set rng = Range("B:B")


first_day = Application.WorksheetFunction.Min(Range("B:B"))
last_day = Application.WorksheetFunction.Max(Range("B:B"))

Dim summary_table_row As LongLong
summary_table_row = 2

Dim lastrow As LongLong


For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row




For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    ticker = ws.Cells(i, 1).Value

    volume = volume + ws.Cells(i, 7).Value

    ws.Range("K" & summary_table_row).Value = ticker
    ws.Range("N" & summary_table_row).Value = volume

        If ws.Cells(i, 2).Value = first_day Then
        open_day = ws.Cells(i, 3).Value
        ElseIf ws.Cells(i, 2).Value = last_day Then
        close_day = ws.Cells(i, 6).Value
    
        End If
        
        

    If open_day = 0 Then
    
    percent_change = "1"

    Else:
    percent_change = FormatPercent(yearly_change / open_day, 1)
    End If
    
    
    
        
    

    ws.Range("L" & summary_table_row).Value = yearly_change
    ws.Range("M" & summary_table_row).Value = percent_change

    summary_table_row = summary_table_row + 1


    volume = 0
    first_day = 0
    last_day = 0
    
    
    

Else


    volume = volume + ws.Cells(i, 7).Value

        If ws.Cells(i, 2).Value = first_day Then
        open_day = ws.Cells(i, 3).Value
        ElseIf ws.Cells(i, 2).Value = last_day Then
        close_day = ws.Cells(i, 6).Value
    
        End If
    
End If

Next i

Next ws




End Sub






