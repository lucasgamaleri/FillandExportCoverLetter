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

' Get the current document
Dim oDoc As Object
oDoc = ThisComponent

' Check if this is a Writer document
If Not oDoc.supportsService("com.sun.star.text.TextDocument") Then
    MsgBox "This macro only works with Writer documents"
    Exit Sub
End If

' Use the searchAndReplace interface
Dim oSearchDesc As Object
Dim oFound As Object

' Replace [COMPANY_NAME]
oSearchDesc = oDoc.createSearchDescriptor()
oSearchDesc.SearchString = "[COMPANY_NAME]"
oSearchDesc.SearchCaseSensitive = True
oFound = oDoc.findFirst(oSearchDesc)
Do While Not IsNull(oFound)
    oFound.String = companyName
    oFound = oDoc.findNext(oFound.End, oSearchDesc)
Loop

' Replace [CITY_ADDRESS]
oSearchDesc.SearchString = "[CITY_ADDRESS]"
oFound = oDoc.findFirst(oSearchDesc)
Do While Not IsNull(oFound)
    oFound.String = cityAddress
    oFound = oDoc.findNext(oFound.End, oSearchDesc)
Loop

' Replace [COUNTRY]
oSearchDesc.SearchString = "[COUNTRY]"
oFound = oDoc.findFirst(oSearchDesc)
Do While Not IsNull(oFound)
    oFound.String = countryName
    oFound = oDoc.findNext(oFound.End, oSearchDesc)
Loop

' Replace [POSITION_NAME]
oSearchDesc.SearchString = "[POSITION_NAME]"
oFound = oDoc.findFirst(oSearchDesc)
Do While Not IsNull(oFound)
    oFound.String = positionName
    oFound = oDoc.findNext(oFound.End, oSearchDesc)
Loop

MsgBox "Cover letter has been updated successfully"
End Sub

' Exporting to PDF format
Sub ExportToPDF()
    Dim oDoc As Object
    Dim docURL As String
    Dim sFilter As String
    Dim aProps(0) As New com.sun.star.beans.PropertyValue
    
    ' Get current document
    oDoc = ThisComponent
    
    ' Check if document has been saved
    If oDoc.hasLocation() = False Then
        MsgBox "Please save the document first before exporting to PDF"
        Exit Sub
    End If
    
    docURL = oDoc.getURL()
    
    ' Create PDF URL by replacing extension
    Dim pdfURL As String
    Dim lastDot As Integer
    lastDot = InStrRev(docURL, ".")
    
    If lastDot > 0 Then
        pdfURL = Left(docURL, lastDot - 1) & ".pdf"
    Else
        pdfURL = docURL & ".pdf"
    End If
    
    ' Set up PDF export filter
    aProps(0).Name = "FilterName"
    aProps(0).Value = "writer_pdf_Export"
    
    ' Export to PDF
    On Error GoTo ErrorHandler
    oDoc.storeToURL(pdfURL, aProps())
    
    MsgBox "Document exported to PDF successfully!"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error exporting to PDF: " & Error$ & " (Error " & Err & ")"
End Sub

Sub FillAndExportCoverLetter()
' Combined macro that fills and exports to PDF
Call FillCoverLetterFields
Call ExportToPDF
End Sub

REM  *****  BASIC  *****
