'Runs automatically when opening workbook.xlsm

Sub Auto_Open()
    MsgBox "Welcome!"
End Sub

'Runs automatically when closing workbook.xlsm, checking and saving this workbook

Sub Auto_Close()
    If ThisWorkbook.Saved = False Then
        ThisWorkbook.Save
    End If
End Sub


'VBA macros can be written within a Sub or a Function

'Function returns a result and can be invoked as a formula, is used for calculations
'It is not shown in "Macros" section in Excel

Function User()
    User = Application.UserName
End Function

'Sub does not return a value to use, doing what is programmed instead
'For example, writing values in cells

Sub CellReferencing()
    ActiveCell.Value = "Value"
    [B1].Value = "B1"
End Sub


