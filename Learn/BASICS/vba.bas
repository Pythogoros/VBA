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


'Ranges can be used instead of individual cells - ALWAYS BETTER TO SPECIFY!

Sub Bold()
    [C1:D3].Font.Bold = True
    Range("E4:F6").Font.Italic = True
End Sub


'Sheets can be accessed in different ways - ALWAYS BETTER TO SPECIFY!

Sub Sheets()
    ActiveSheet.[A1].Value = "A1"
    Sheets("Sheet1").[A1].Font.Bold = True
End Sub


'Variables can be defined for latter usage

Sub Italic()
    Dim Range1 as Range
    Set Range1 = Sheets("Sheet1").Range("A1:B3")
    Range1.Font.Italic = True
End Sub


'Cells stand for all the cells in the current sheet

Sub Cells()
    Cells.EntireColumn.Autofit
    Cells.EntireRow.Autofit
    Cells.WrapText = True
End Sub


'User inputs are allowed

Sub UserInput()
    Dim sheets2add as Integer
    sheets2add = InputBox("Enter number of sheets to add:")
    Sheets.Add Count:=sheets2add, Before:=ActiveSheet
End Sub



'Can insert columns with user input

Sub InsertColumns()

    Dim cols as Integer
    Dim nums as Integer

    ActiveCell.EntireColumn.Select
    cols = InputBox("Enter number of columns to add:")
    For nums = 1 to cols
        Selection.Insert Shift:=xlToLeft
    Next nums

End Sub


'Can also insert rows with user input

Sub InsertRows()

    Dim rows as Integer
    Dim nums as Integer

    ActiveCell.EntireRow.Select
    rows = InputBox("Enter number of rows to add:")
    For nums = 1 to rows
        Selection.Insert Shift:=xlToUp
    Next rows

End Sub


'Can merge and unmerge cells

Sub MergeCells()
    Dim Range as Range
    Set Range = Range("A1:B2")
    Range.Merge
End Sub

Sub UnMergeCells()
    Dim Range as Range
    Set Range = Range("A1:B2")
    Range.UnMerge
End Sub


'Can hide and unhide cells

Sub HideColumns()

    Dim Range as Range

    For Each Range in Selection
        Range.EntireColumn.Hidden = True
    Next Range

End Sub


Sub UnHideColumns()

    Dim Range as Range

    For Each Range in Selection
        Range.EntireColumn.Hidden = False
    Next Range

End Sub


Sub HideRows()

    Dim Range as Range

    For Each Range in Selection
        Range.EntireRow.Hidden = True
    Next Range

End Sub


Sub UnHideRows()

    Dim Range as Range

    For Each Range in Selection
        Range.EntireRow.Hidden = False
    Next Range

End Sub


'Can generate serial numbers from 1 to user input

Sub SN()

    Dim sn as Integer

    sn =  InputBox("Enter number of values to add:")
    For sn = 1 to sn
        ActiveCell.Value = sn
        ActiveCell.Offset(1,0).Activate
    Next sn

End Sub


'Can do Dim outside Sub() to use later anywhere

Dim Range as Range

Sub Sub1()
    Range...
End Sub

Sub Sub2()
    Range...
End Sub


'Can select wide ranges simulating ctrl+arrow

Sub WR()

    Dim Range as Range

    Set Range = Range([A1], [A1].End(xlDown))

    ...
    
End Sub


'Can do conditional formatting and clear it

Sub CF()

    Dim Range as Range
    Dim a as FormatCondition
    Dim b as FormatCondition

    Set a = Range.FormatConditions.Add(xlCellValue, xlGreater, 0)
    Set b = Range.FormatConditions.Add(xlCellValue, xlLess, 0)

    a.Interior.Color = RGB(0, 255, 0)
    b.Interior.Color = RGB(255, 0, 0)

    Range.FormatConditions.Delete 'option 1
    Cells.FormatConditions.Delete 'option 2

End Sub


'Can even insert formulas

Sub Formula()

    ...
        [A1].FormulaR1C1 = "=R[1]C[1]/R[2]C[2]"
    ...

End Sub