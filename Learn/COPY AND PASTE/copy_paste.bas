Sub CP()

    Application.ScreenUpdating = False

    Dim main As Worksheet
    Set main = ThisWorkbook.Sheets("Main")

    Dim pnl As Worksheet
    Set pnl = ThisWorkbook.Sheets("PnL")

    Dim macro As Worksheet
    Set macro = ThisWorkbook.Sheets("Macro")

    For i = 1 To macro.[R2].Value

        main.[C6] = macro.[R4].Cells(i, 1).Value

        pnl.[G11].Copy
        macro.[S4].Cells(i, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    Next i

    Application.ScreenUpdating = True

    ThisWorkbook.Save

End Sub