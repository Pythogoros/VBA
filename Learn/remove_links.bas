Sub RemoveAllLinks()

    Dim wb As Workbook
    Dim i As Long
    Dim links As Variant

    Set wb = ThisWorkbook

    On Error Resume Next
    links = wb.LinkSources(xlExcelLinks)
    On Error GoTo 0

    If IsEmpty(links) Then
        MsgBox "Not Found", vbInformation
        Exit Sub
    End If

    For i = UBound(links) To LBound(links) Step -1
        wb.BreakLink Name:=links(i), Type:=xlExcelLinks
    Next i
    
End Sub