Attribute VB_Name = "DocGenerator"
Option Explicit

Const WD_YELLOW As Long = 7
Const WD_NO_HIGHLIGHT As Long = 0

'===========================================================
' Сканирование шаблонов и построение/обновление листов реестра
'===========================================================
Public Sub ScanTemplates()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim templatePath As String
    templatePath = wb.Path & "\Шаблоны\"
    
    If Dir(templatePath, vbDirectory) = "" Then
        MsgBox "Не найдена папка шаблонов:" & vbCrLf & templatePath, vbExclamation, "Шаблоны"
        Exit Sub
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(templatePath)
    
    Dim wdApp As Object
    Dim wasOpen As Boolean
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    wasOpen = (Err.Number = 0)
    On Error GoTo 0
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    
    If wdApp Is Nothing Then
        MsgBox "Не удалось подключиться к Microsoft Word." & vbCrLf & "Убедитесь, что Word установлен.", vbCritical
        Exit Sub
    End If
    
    wdApp.Visible = False
    wdApp.DisplayAlerts = False
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        Dim sName As String
        sName = subFolder.Name
        
        ' Пропуск скрытых/системных
        If Left(sName, 1) = "." Then GoTo SkipFolder
        If Left(sName, 1) = "~" Then GoTo SkipFolder
        
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = wb.Worksheets(sName)
        On Error GoTo 0
        
        If ws Is Nothing Then
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = sName
            ws.Tab.Color = RGB(255, 255, 0)
        End If
        
        ' Удаляем старые заголовки (кроме данных пользователя)
        ' Просто очищаем первую строку и перестраиваем
        ws.Rows(1).ClearContents
        ws.Rows(1).ClearFormats
        
        ' Собираем уникальные жёлтые фразы из всех docx в подпапке
        Dim dictPhrases As Object
        Set dictPhrases = CreateObject("Scripting.Dictionary")
        
        Dim f As Object
        For Each f In subFolder.Files
            If LCase(fso.GetExtensionName(f.Name)) = "docx" Then
                Dim doc As Object
                Set doc = wdApp.Documents.Open(f.Path, ReadOnly:=True, Visible:=False)
                
                Dim r As Object
                Set r = doc.Content
                With r.Find
                    .ClearFormatting
                    .Highlight = True
                    .Forward = True
                    .Wrap = 0
                    Do While .Execute
                        If r.HighlightColorIndex = WD_YELLOW Then
                            Dim txt As String
                            txt = Trim(r.Text)
                            txt = Replace(txt, vbCr, " ")
                            txt = Replace(txt, vbLf, " ")
                            txt = Replace(txt, vbCrLf, " ")
                            txt = Application.Trim(txt)
                            If Len(txt) > 0 Then
                                If Not dictPhrases.Exists(txt) Then
                                    dictPhrases.Add txt, txt
                                End If
                            End If
                        End If
                        r.Collapse Direction:=0
                    Loop
                End With
                
                doc.Close SaveChanges:=False
                Set doc = Nothing
            End If
        Next f
        
        ' Записываем заголовки в первую строку
        Dim idx As Long
        idx = 1
        Dim k As Variant
        For Each k In dictPhrases.Keys
            ws.Cells(1, idx).Value = CStr(k)
            idx = idx + 1
        Next k
        
        ' Оформление заголовков
        If idx > 1 Then
            Dim lastCol As Long
            lastCol = idx - 1
            With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
                .Font.Bold = True
                .Interior.Color = RGB(200, 200, 200)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            ws.Rows(1).AutoFit
        End If
        
        Set ws = Nothing
SkipFolder:
    Next subFolder
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    If Not wasOpen Then wdApp.Quit
    Set wdApp = Nothing
    
    MsgBox "Реестр обновлён из шаблонов." & vbCrLf & "Заполните строки и нажмите «Сгенерировать документы».", vbInformation
End Sub

'===========================================================
' Генерация заполненных документов
'===========================================================
Public Sub GenerateDocuments()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim templatePath As String
    templatePath = wb.Path & "\Шаблоны\"
    Dim outputPath As String
    outputPath = wb.Path & "\Выход\"
    
    If Dir(templatePath, vbDirectory) = "" Then
        MsgBox "Папка шаблонов не найдена!", vbExclamation
        Exit Sub
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Dir(outputPath, vbDirectory) = "" Then MkDir outputPath
    
    Dim wdApp As Object
    Dim wasOpen2 As Boolean
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    wasOpen2 = (Err.Number = 0)
    On Error GoTo 0
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    
    If wdApp Is Nothing Then
        MsgBox "Не удалось подключиться к Microsoft Word!", vbCritical
        Exit Sub
    End If
    
    wdApp.Visible = False
    wdApp.DisplayAlerts = False
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Генерация документов..."
    
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim checkPath As String
        checkPath = templatePath & ws.Name & "\"
        If Dir(checkPath, vbDirectory) = "" Then GoTo NextWs
        
        Dim lRow As Long
        lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lRow < 2 Then GoTo NextWs
        
        ' Заголовки в словарь
        Dim dictH As Object
        Set dictH = CreateObject("Scripting.Dictionary")
        Dim lCol As Long
        lCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Dim c As Long
        For c = 1 To lCol
            Dim hv As String
            hv = Trim(ws.Cells(1, c).Value)
            If Len(hv) > 0 And Not dictH.Exists(hv) Then
                dictH.Add hv, c
            End If
        Next c
        
        Dim r As Long
        For r = 2 To lRow
            ' Пропуск полностью пустых строк
            If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, lCol))) = 0 Then GoTo NextRow
            
            Dim sheetOutPath As String
            sheetOutPath = outputPath & ws.Name & "\"
            If Dir(sheetOutPath, vbDirectory) = "" Then MkDir sheetOutPath
            
            Dim rowFolder As String
            rowFolder = sheetOutPath & "Строка_" & Format(r, "000") & "\"
            If Dir(rowFolder, vbDirectory) = "" Then MkDir rowFolder
            
            Dim srcFolder As Object
            Set srcFolder = fso.GetFolder(checkPath)
            Dim fl As Object
            For Each fl In srcFolder.Files
                If LCase(fso.GetExtensionName(fl.Name)) = "docx" Then
                    Dim outFile As String
                    outFile = rowFolder & fl.Name
                    fso.CopyFile fl.Path, outFile, True
                    
                    Dim doc As Object
                    Set doc = wdApp.Documents.Open(outFile, ReadOnly:=False, Visible:=False)
                    
                    Dim rngw As Object
                    Set rngw = doc.Content
                    With rngw.Find
                        .ClearFormatting
                        .Highlight = True
                        .Forward = True
                        .Wrap = 0
                        Do While .Execute
                            If rngw.HighlightColorIndex = WD_YELLOW Then
                                Dim ph As String
                                ph = Trim(rngw.Text)
                                ph = Replace(ph, vbCr, " ")
                                ph = Replace(ph, vbLf, " ")
                                ph = Replace(ph, vbCrLf, " ")
                                ph = Application.Trim(ph)
                                
                                If dictH.Exists(ph) Then
                                    Dim ci As Long
                                    ci = dictH(ph)
                                    Dim rv As String
                                    rv = Trim(ws.Cells(r, ci).Value & "")
                                    
                                    rngw.Text = rv
                                    rngw.HighlightColorIndex = WD_NO_HIGHLIGHT
                                End If
                            End If
                            rngw.Collapse Direction:=0
                        Loop
                    End With
                    
                    doc.Close SaveChanges:=True
                    Set doc = Nothing
                End If
            Next fl
NextRow:
        Next r
NextWs:
    Next ws
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    If Not wasOpen2 Then wdApp.Quit
    Set wdApp = Nothing
    
    MsgBox "Генерация завершена!" & vbCrLf & "Документы сохранены в папке:" & vbCrLf & outputPath, vbInformation
End Sub
