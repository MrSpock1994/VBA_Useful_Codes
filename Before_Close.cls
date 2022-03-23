Private Sub Workbook_BeforeClose(Cancel As Boolean)

'Simple macro to stop the user from closing the workbook if a certain condition is not meet


Dim wkOrigem As Worksheet

Set wkDestino = Workbooks("VBA_Code.xlsm").Worksheets("Testing")

' For this case the condition is if the cell(1, 1) is empty the user can not close the workbook

With wkDestino
verificalcel = ThisWorkbook.Sheets("Testing").Cells(1, 1).Value
    If verificalcel = "" Then

        Cancel = True
        MsgBox "ERRO"
    Else

        Cancel = False

    End If

End With

End Sub

