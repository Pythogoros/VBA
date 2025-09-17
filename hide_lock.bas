Sub Hide_and_Lock()

    Dim wb As Workbook
    Set wb = ThisWorkbook


    'Hide

    For Each sh In wb.Worksheets
    If sh.Name <> "Sheet1" Then sh.Visible = False
    End If
    Next


    'Call another sub to remove links

    Call RemoveAllLinks


    'Lock

    wb.Protect Password:="Password", Structure:=True, Windows:=False
    wb.Sheets("Sheet1").Protect Password:="Password", DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowFormattingRows:=False, AllowUsingPivotTables:=True

    filepath_closed = "PATH:\path\"
    Filename = "FILE"
    filesave = filepath_closed & Filename & ".xlsb"
    wb.SaveAs Filename:=filesave

End Sub