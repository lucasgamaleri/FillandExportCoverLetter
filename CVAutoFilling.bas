REM  *****  BASIC  *****
Sub FillCoverLetterFields()
'
' FillCoverLetterFields Macro
'
'

' Declare variables for company information

Dim companyName As String
Dim cityAddress As String
Dim positionName As String
Dim countryName As String

' Get company information
companyName = InputBox("Enter company Name", "Company info", "ABC company")
cityAddress = InputBox("Enter address, city:", "Company info", "456 Business Address")
countryName = InputBox("Enter country name:", "Company info", "Australia")
positionName = InputBox("Enter position name:", "Job info", "Mechanical Engineer")

' Replace placeholders with actual information
With ActiveDocument.Range.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    
    ' Replace information
    .Text = "[COMPANY_NAME]"
    .Replacement.Text = companyName
    .Execute Replace:=wdReplaceAll
    
    .Text = "[CITY_ADDRESS]"
    .Replacement.Text = cityAddress
    .Execute Replace:=wdReplaceAll
    
    .Text = "[COUNTRY]"
    .Replacement.Text = countryName
    .Execute Replace:=wdReplaceAll
    
    .Text = "[POSITION_NAME]"
    .Replacement.Text = positionName
    .Execute Replace:=wdReplaceAll
End With

MsgBox "Cover letter has been updated successfully"

End Sub


' Exporting to PDF format

Sub ExportToPDF()

    ' Get current document name without extension
    Dim docName As String
    Dim pdfPath As String
    
    docName = Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1)
    
    ' Create PDF filename with timestamp
    pdfPath = ActiveDocument.Path & "\" & docName & Format(Now, "yyyymmdd_hhmm") & ".pdf"
    
    ' Export to PDF
    ActiveDocument.ExportAsFixedFormat _
    OutputFileName:=pdfPath, _
    ExportFormat:=wdExportFormatPDF, _
    OpenAfterExport:=False, _
    OptimizeFor:=wdExportOptimizeForMinimumSize, _
    BitMapMissingFonts:=True, _
    DocStructureTags:=False, _
    CreateBookmarks:=wdExportCreateNoBookmarks
    
    

    
MsgBox "Document exported to PDF: " & pdfPath

End Sub

Sub FillandExportCoverLetter()

' Combined macro that fills and export to pdf

Call FillCoverLetterFields
Call ExportToPDF

End Sub



