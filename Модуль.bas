Option Explicit

Sub GeneratePackage()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim ws As Worksheet
    Dim templateBasePath As String
    Dim outputBasePath As String
    
    templateBasePath = wb.Path & "\Шаблоны\"
    outputBasePath = wb.Path & "\Выход\"
    
    ' Ensure output base folder exists
    If Dir(outputBasePath, vbDirectory) = "" Then MkDir outputBasePath
    
    ' Loop through each sheet (each sheet is a template set)
    For Each ws In wb.Worksheets
        ' Skip hidden or very hidden sheets if needed
        If ws.Visible = xlSheetVisible Then
            ProcessSheet ws, templateBasePath, outputBasePath
        End If
    Next ws
    
    MsgBox "Пакеты сформированы!", vbInformation
End Sub

Sub ProcessSheet(ws As Worksheet, templateBasePath As String, outputBasePath As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim headerRow As Range
    Set headerRow = ws.Rows(1)
    
    Dim i As Long
    For i = 2 To lastRow
        Dim contractNumber As String
        contractNumber = Trim(ws.Cells(i, GetColumnIndex(ws, "номер договора")).Value)
        If contractNumber = "" Then GoTo NextRow
        
        Dim outputFolder As String
        outputFolder = outputBasePath & contractNumber & "\"
        If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder
        
        ' Process each template type
        Dim templateTypes As Variant
        templateTypes = Array("Договор", "Заявление", "Прилож3", "Счёт", "Лист")
        
        Dim t As Variant
        For Each t In templateTypes
            Dim templateFolder As String
            templateFolder = templateBasePath & t & "\"
            
            ' Determine the set number from the sheet name (e.g., Набор1 -> 1)
            Dim setNumber As String
            setNumber = Replace(ws.Name, "Набор", "")
            If Not IsNumeric(setNumber) Then setNumber = "1"
            
            Dim templateFile As String
            templateFile = templateFolder & setNumber & ". " & t & ".docx"
            
            If Dir(templateFile) <> "" Then
                Dim outputFile As String
                outputFile = outputFolder & t & ".docx"
                
                CopyAndReplace templateFile, outputFile, ws, i
            End If
        Next t
NextRow:
    Next i
End Sub

Sub CopyAndReplace(templatePath As String, outputPath As String, ws As Worksheet, rowIndex As Long)
    On Error GoTo ErrHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.CopyFile templatePath, outputPath, True
    
    Dim wdApp As Object
    Dim wdDoc As Object
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    
    Set wdDoc = wdApp.Documents.Open(outputPath, ReadOnly:=False)
    
    Dim placeholder As String
    Dim replacement As String
    
    Dim colIndex As Long
    For colIndex = 1 To ws.Columns.Count
        Dim headerValue As String
        headerValue = Trim(ws.Cells(1, colIndex).Value)
        If headerValue = "" Then GoTo NextCol
        
        placeholder = "{{" & headerValue & "}}"
        replacement = Trim(ws.Cells(rowIndex, colIndex).Value)
        
        If replacement = "" Then replacement = ""  ' keep empty
        
        With wdDoc.Content.Find
            .Text = placeholder
            .Replacement.Text = replacement
            .Forward = True
            .Wrap = 1  ' wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=2  ' wdReplaceAll
        End With
NextCol:
    Next colIndex
    
    wdDoc.SaveAs2 outputFileName:=outputPath, FileFormat:=wdFormatXMLDocument
    wdDoc.Close
    wdApp.Quit
    
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "Error processing " & templatePath & ": " & Err.Description, vbCritical
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close
    If Not wdApp Is Nothing Then wdApp.Quit
End Sub

' Helper to get column index by header name (case-insensitive)
Function GetColumnIndex(ws As Worksheet, headerName As String) As Long
    Dim c As Range
    For Each c In ws.Rows(1).Cells
        If Trim(LCase(c.Value)) = Trim(LCase(headerName)) Then
            GetColumnIndex = c.Column
            Exit Function
        End If
    Next c
    GetColumnIndex = 0  ' not found
End Function