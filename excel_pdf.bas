Sub Excel_and_PDF()


    'Variables

    Dim Filesave As String
    Dim Filename As String

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim DataRng As Range
    Dim strfile As String
    Dim pdfile As String


    'Setting Workbook

    Set wb = ThisWorkbook
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set DataRng = ws.Range("A1")


    'Setting Paths

    Filepath_excl = "PATH:\"
    Filepath_pdf = "PATH:\path\"


    'Saving Workbook with given name and current time

    Filename = ws.Range("A1").Value
    CurrentTime = Format(ws.Range("B1").Value, "yyyymmdd_HHmm")
    Filesave = Filepath_excl & Filename & "_" & CurrentTime & ".xlsm"
    wb.SaveAs Filename:=Filesave
    pdfname = Filename & " " & CurrentTime & ".pdf"
    pdfile = Filepath_pdf & pdfname
    DataRng.ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=pdfile, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=True

End Sub