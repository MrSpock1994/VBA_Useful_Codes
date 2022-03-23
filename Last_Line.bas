Attribute VB_Name = "Módulo1"
Sub Last_line()

Dim x As Integer
Dim wkOrigem As Worksheet

Set wkOrigem = Workbooks("VBA_Code.xlsm").Worksheets("Testing")

'Macro to find the last cell with a value

x = 1
With wkOrigem

check_last = ThisWorkbook.Sheets("Testing").Cells(x, 1).Value

Do While check_last <> ""
    x = x + 1
    check_last = ThisWorkbook.Sheets("Testing").Cells(x, 1).Value
    
Loop

'x is the line with blank value, so x-1 is the last line with a value

MsgBox (x - 1)
    

End With


End Sub
