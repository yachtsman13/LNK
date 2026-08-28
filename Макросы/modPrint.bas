Attribute VB_Name = "modPrint"
Option Explicit

' Собирает все сформированные документы одного объекта (совпадение объекта и даты
' в листе "Журнал") в один Word-файл: сквозная нумерация без номера на 1-й странице,
' пустой лист в конце под сшив, печать в 3 экземплярах.
'
' ВАЖНО (реальное ограничение, а не недоработка макроса): свидетельства о поверке
' и квалификационные удостоверения — отдельные сканы/PDF, а не Word-файлы. Только
' Adobe Reader (без Acrobat) не даёт программно склеить PDF и Word в один пронумерованный
' файл без прав администратора на установку ПО. Эти страницы печатаются из Adobe Reader
' отдельно и подкладываются в пакет вручную в нужном месте — так и должно быть с этим
' набором программ, макрос эту часть не подменяет.
Sub СобратьИНапечататьПакет()
    Dim ws As Worksheet, r As Long, lastRow As Long, baseRow As Long
    Dim objName As String, custName As String, ctrlDate As Variant
    Dim wdApp As Object, masterDoc As Object
    Dim files As Collection
    Dim docPath As String
    Dim i As Long
    Dim outFolder As String, outName As String

    Set ws = ActiveSheet
    If ws.Name <> "Журнал" Then
        MsgBox "Встаньте на строку объекта в листе ""Журнал"" и запустите макрос ещё раз.", vbExclamation, "Сборка пакета"
        Exit Sub
    End If

    baseRow = ActiveCell.Row
    objName = Trim(ws.Cells(baseRow, COL_OBJECT).Value)
    ctrlDate = ws.Cells(baseRow, COL_DATE).Value
    custName = Trim(ws.Cells(baseRow, COL_CUSTOMER).Value)

    If objName = "" Then
        MsgBox "В выбранной строке не заполнен объект контроля.", vbExclamation, "Сборка пакета"
        Exit Sub
    End If

    Set files = New Collection
    lastRow = ws.Cells(ws.Rows.Count, COL_METHOD).End(xlUp).Row
    For r = 5 To lastRow
        If Trim(ws.Cells(r, COL_OBJECT).Value) = objName _
           And CStr(ws.Cells(r, COL_DATE).Value) = CStr(ctrlDate) _
           And ws.Cells(r, COL_GENERATED).Value <> "" Then
            docPath = ПутьКДокументу(ws, r)
            If Dir(docPath) <> "" Then files.Add docPath
        End If
    Next r

    If files.Count = 0 Then
        MsgBox "Не нашлось сформированных документов по этому объекту и дате." & vbCrLf & _
               "Сначала выполните ""Сформировать документ"" для каждой строки этого объекта.", _
               vbExclamation, "Сборка пакета"
        Exit Sub
    End If

    If MsgBox("Найдено документов для сборки: " & files.Count & vbCrLf & _
              "Объект: " & objName & vbCrLf & vbCrLf & _
              "Собрать в один файл, добавить лист под сшив и отправить на печать (3 экз., двусторонняя)?", _
              vbYesNo + vbQuestion, "Сборка пакета") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set masterDoc = wdApp.Documents.Add

    Dim insertPoint As Object
    For i = 1 To files.Count
        If i > 1 Then
            Set insertPoint = masterDoc.Range(masterDoc.Content.End - 1, masterDoc.Content.End - 1)
            insertPoint.InsertBreak Type:=7   ' wdPageBreak
        End If
        Set insertPoint = masterDoc.Range(masterDoc.Content.End - 1, masterDoc.Content.End - 1)
        insertPoint.InsertFile files(i)
    Next i

    ' Пустой лист в конце — под бирку сшива
    Set insertPoint = masterDoc.Range(masterDoc.Content.End - 1, masterDoc.Content.End - 1)
    insertPoint.InsertBreak Type:=7

    ' Сквозная нумерация без номера на первой странице (весь файл — одна секция)
    masterDoc.Sections(1).PageSetup.DifferentFirstPageHeaderFooter = True
    masterDoc.Sections(1).Footers(1).PageNumbers.Add PageNumberAlignment:=1, FirstPage:=False  ' wdAlignPageNumberCenter

    outFolder = OutputFolder() & ws.Cells(baseRow, COL_YEAR).Value & "\Пакеты\"
    MkPath outFolder
    outName = outFolder & ОчиститьДляИмениФайла(objName) & " " & Format(ctrlDate, "dd-mm-yyyy") & ".docx"
    masterDoc.SaveAs2 fileName:=outName

    masterDoc.PrintOut Copies:=3, Background:=False

    Application.ScreenUpdating = True

    MsgBox "Пакет собран и отправлен на печать (3 экз.):" & vbCrLf & outName & vbCrLf & vbCrLf & _
           "Свидетельства о поверке и удостоверения специалистов — отдельные сканы, печатаются " & _
           "из Adobe Reader и подкладываются в пакет вручную по месту. Прошивку и бирку — по правилам " & _
           "делопроизводства (см. лист ""Меню"" → шаг 7).", _
           vbInformation, "Готово"

    masterDoc.Close SaveChanges:=False
    wdApp.Quit
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not masterDoc Is Nothing Then masterDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    MsgBox "Не удалось собрать пакет." & vbCrLf & "Ошибка: " & Err.Description & vbCrLf & vbCrLf & _
           "Проверьте на 2-3 документах перед тем, как полагаться на этот макрос в реальной работе — " & _
           "он не был протестирован в настоящем Word.", vbCritical, "Сборка пакета"
End Sub

' Пакетная печать всех документов одного метода из папки "Готовые" —
' исправленная версия старого макроса ПЕЧАТЬ_ПРОТОКОЛОВ (в исходнике было
' два одноимённых Sub в разных модулях + функция, объявленная внутри Sub —
' из-за этого VBA-проект вообще не компилировался).
Sub ПечатьПоМетоду()
    Dim methodName As String
    Dim folder As String
    Dim fso As Object, subf As Object, ff As Object
    Dim wdApp As Object, wdDoc As Object
    Dim printed As Long

    methodName = InputBox("Метод контроля для печати (ВИК, УЗТ, УЗД, ТВ, АЭ, ТК, ВД):", "Пакетная печать")
    If Trim(methodName) = "" Then Exit Sub

    folder = OutputFolder()
    If Dir(folder, vbDirectory) = "" Then
        MsgBox "Папка ""Готовые"" ещё не создана — сначала сформируйте хотя бы один документ.", vbExclamation
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    printed = 0

    On Error GoTo CleanFail
    Dim yearFolder As Object
    For Each yearFolder In fso.GetFolder(folder).SubFolders
        Dim methodFolder As String
        methodFolder = yearFolder.Path & "\" & methodName
        If fso.FolderExists(methodFolder) Then
            For Each ff In fso.GetFolder(methodFolder).Files
                If LCase(fso.GetExtensionName(ff.Name)) = "docx" Then
                    Set wdDoc = wdApp.Documents.Open(ff.Path, ReadOnly:=True, AddToRecentFiles:=False)
                    wdDoc.PrintOut Background:=False
                    wdDoc.Close SaveChanges:=False
                    Set wdDoc = Nothing
                    printed = printed + 1
                End If
            Next ff
        End If
    Next yearFolder

    wdApp.Quit
    MsgBox "Отправлено на печать документов: " & printed, vbInformation, "Пакетная печать"
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    MsgBox "Ошибка при пакетной печати: " & Err.Description, vbCritical, "Пакетная печать"
End Sub
