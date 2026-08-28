Attribute VB_Name = "modGenerate"
Option Explicit

' Формирует документ (Акт/Протокол/Заключение) для строки, на которой стоит курсор
' на листе "Журнал". Подставляет то, что реально хранится в Журнале и реестрах;
' остальные поля (объём, средства контроля, результаты измерений, применённые НД)
' остаются плейсхолдерами #ИмяПоля для ручного заполнения в открывшемся документе —
' этих данных в Журнале нет, придумывать их макрос не должен.
Sub СформироватьДокумент()
    Dim ws As Worksheet
    Dim r As Long
    Dim methodName As String, regNum As String, objName As String
    Dim custName As String, execName As String
    Dim ctrlDate As Variant
    Dim tplFile As String, tplPath As String, outPath As String, outFolder As String
    Dim wdApp As Object, wdDoc As Object
    Dim warnMsg As String

    Set ws = ActiveSheet
    If ws.Name <> "Журнал" Then
        MsgBox "Встаньте на нужную строку листа ""Журнал"" и запустите макрос ещё раз.", vbExclamation, "Формирование документа"
        Exit Sub
    End If

    r = ActiveCell.Row
    If r < 5 Then
        MsgBox "Выберите строку с данными (начиная со строки 5).", vbExclamation, "Формирование документа"
        Exit Sub
    End If

    methodName = Trim(ws.Cells(r, COL_METHOD).Value)
    If methodName = "" Then
        MsgBox "В этой строке не указан метод контроля.", vbExclamation, "Формирование документа"
        Exit Sub
    End If

    tplFile = TemplateFileFor(methodName)
    If tplFile = "" Then
        MsgBox "Для метода """ & methodName & """ пока нет шаблона документа в папке ""Шаблоны"".", _
               vbExclamation, "Формирование документа"
        Exit Sub
    End If

    tplPath = TemplatesFolder() & tplFile
    If Dir(tplPath) = "" Then
        MsgBox "Не найден файл шаблона:" & vbCrLf & tplPath & vbCrLf & vbCrLf & _
               "Проверьте, что папка ""Шаблоны"" лежит рядом с этой книгой Excel.", _
               vbCritical, "Формирование документа"
        Exit Sub
    End If

    regNum = ws.Cells(r, COL_REGNUM).Text
    objName = Trim(ws.Cells(r, COL_OBJECT).Value)
    custName = Trim(ws.Cells(r, COL_CUSTOMER).Value)
    execName = Trim(ws.Cells(r, COL_EXECUTOR).Value)
    ctrlDate = ws.Cells(r, COL_DATE).Value

    If objName = "" Then
        If MsgBox("В строке не заполнен объект контроля. Сформировать документ всё равно?", _
                  vbYesNo + vbQuestion, "Формирование документа") = vbNo Then Exit Sub
    End If

    ' --- Проверка сроков аттестации исполнителя ПЕРЕД формированием ---
    If execName <> "" Then
        warnMsg = ПроверитьАттестацию(execName, methodName, ctrlDate)
        If warnMsg <> "" Then
            If MsgBox(warnMsg & vbCrLf & vbCrLf & "Всё равно сформировать документ?", _
                      vbYesNo + vbExclamation, "Проверка сроков") = vbNo Then
                Exit Sub
            End If
        End If
    End If

    If ws.Cells(r, COL_GENERATED).Value <> "" Then
        If MsgBox("Документ по этой строке уже формировался " & _
                  Format(ws.Cells(r, COL_GENERATED).Value, "dd.mm.yyyy hh:mm") & _
                  ". Сформировать заново (старый файл будет перезаписан)?", _
                  vbYesNo + vbQuestion, "Формирование документа") = vbNo Then
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(fileName:=tplPath, ReadOnly:=False, AddToRecentFiles:=False)

    ЗаменитьПлейсхолдер wdDoc, "Дата_и_место_проведения_контроля", ДатаМесто(ctrlDate, custName)
    ЗаменитьПлейсхолдер wdDoc, "Номер", CStr(ws.Cells(r, COL_SERIES).Value)
    ЗаменитьПлейсхолдер wdDoc, "Год", CStr(ws.Cells(r, COL_YEAR).Value)
    ЗаменитьПлейсхолдер wdDoc, "Объект_контроля", objName
    ЗаменитьПлейсхолдер wdDoc, "Наименование_объекта", objName
    ЗаменитьПлейсхолдер wdDoc, "Заказывающее_подразделение", custName
    ЗаменитьПлейсхолдер wdDoc, "Уровень_квалификации_№_удостоверения_ФИО", execName
    ЗаменитьПлейсхолдер wdDoc, "Уровень_№_ФИО", execName

    outFolder = OutputFolder() & ws.Cells(r, COL_YEAR).Value & "\" & methodName & "\"
    MkPath outFolder
    outPath = ПутьКДокументу(ws, r)

    wdDoc.SaveAs2 fileName:=outPath
    wdDoc.Close SaveChanges:=False
    wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing

    Application.ScreenUpdating = True

    ws.Cells(r, COL_GENERATED).Value = Now
    ws.Cells(r, COL_GENERATED).NumberFormat = "dd.mm.yyyy hh:mm"

    MsgBox "Документ сформирован:" & vbCrLf & outPath & vbCrLf & vbCrLf & _
           "В Журнале не хранятся: объём/заводской № объекта, средства контроля, результаты " & _
           "измерений, применённые нормативные документы, ФИО руководителя работ — эти поля " & _
           "остались как #Плейсхолдеры, дозаполните их в открывшемся документе вручную.", _
           vbInformation, "Готово"

    On Error Resume Next
    Shell "explorer.exe /select,""" & outPath & """", vbNormalFocus
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    MsgBox "Не удалось сформировать документ." & vbCrLf & "Ошибка: " & Err.Description, vbCritical, "Формирование документа"
End Sub

Private Sub ЗаменитьПлейсхолдер(doc As Object, fieldName As String, value As String)
    Dim rng As Object
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Text = "#" & fieldName
        .Replacement.ClearFormatting
        .Replacement.Text = value
        .Forward = True
        .Wrap = 1        ' wdFindContinue
        .Format = False
        .MatchCase = False
        .Execute Replace:=2   ' wdReplaceAll
    End With
End Sub

Private Function ДатаМесто(d As Variant, cust As String) As String
    Dim s As String
    If IsDate(d) Then s = Format(d, "dd.mm.yyyy") Else s = ""
    If cust <> "" Then
        If s <> "" Then s = s & ", "
        s = s & cust
    End If
    ДатаМесто = s
End Function
