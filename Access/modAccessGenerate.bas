Attribute VB_Name = "modAccessGenerate"
Option Explicit
Option Compare Database

' Вызывается кнопкой "Сформировать документ" на форме Форма_Журнал.
' Читает данные ТЕКУЩЕЙ ЗАПИСИ формы (Me), а не ячейки — в этом разница
' с версией для Excel. Логика подстановки и проверки сроков та же.
Public Sub СформироватьДокумент(frm As Form)
    Dim methodName As String, regNum As String, objName As String
    Dim custName As String, execName As String
    Dim ctrlDate As Variant
    Dim tplFile As String, tplPath As String, outPath As String, outFolder As String
    Dim wdApp As Object, wdDoc As Object
    Dim warnMsg As String

    methodName = Nz(frm!Метод, "")
    If methodName = "" Then
        MsgBox "В этой записи не указан метод контроля.", vbExclamation, "Формирование документа"
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
               "Проверьте, что папка ""Шаблоны"" лежит рядом с базой данных.", vbCritical, "Формирование документа"
        Exit Sub
    End If

    regNum = Nz(frm!РегистрационныйНомер, "")
    objName = Nz(frm!Объект, "")
    custName = Nz(frm!Заказчик, "")
    execName = Nz(frm!Исполнитель, "")
    ctrlDate = frm!ДатаРегистрации

    If objName = "" Then
        If MsgBox("В записи не заполнен объект контроля. Сформировать документ всё равно?", _
                  vbYesNo + vbQuestion, "Формирование документа") = vbNo Then Exit Sub
    End If

    If execName <> "" Then
        warnMsg = ПроверитьАттестацию(execName, methodName, ctrlDate)
        If warnMsg <> "" Then
            If MsgBox(warnMsg & vbCrLf & vbCrLf & "Всё равно сформировать документ?", _
                      vbYesNo + vbExclamation, "Проверка сроков") = vbNo Then
                Exit Sub
            End If
        End If
    End If

    On Error GoTo CleanFail

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(fileName:=tplPath, ReadOnly:=False, AddToRecentFiles:=False)

    ЗаменитьПлейсхолдер wdDoc, "Дата_и_место_проведения_контроля", ДатаМесто(ctrlDate, custName)
    ЗаменитьПлейсхолдер wdDoc, "Номер", CStr(Nz(frm!НомерВСерии, ""))
    ЗаменитьПлейсхолдер wdDoc, "Год", CStr(Nz(frm!Год, ""))
    ЗаменитьПлейсхолдер wdDoc, "Объект_контроля", objName
    ЗаменитьПлейсхолдер wdDoc, "Наименование_объекта", objName
    ЗаменитьПлейсхолдер wdDoc, "Заказывающее_подразделение", custName
    ЗаменитьПлейсхолдер wdDoc, "Уровень_квалификации_№_удостоверения_ФИО", execName
    ЗаменитьПлейсхолдер wdDoc, "Уровень_№_ФИО", execName

    outFolder = OutputFolder() & Nz(frm!Год, "") & "\" & methodName & "\"
    MkPath outFolder
    outPath = outFolder & Replace(regNum, "/", "-") & " " & ОчиститьДляИмениФайла(objName) & ".docx"

    wdDoc.SaveAs2 fileName:=outPath
    wdDoc.Close SaveChanges:=False
    wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing

    frm!ДокументСформирован = Now
    frm.Dirty = False   ' сохранить запись немедленно

    MsgBox "Документ сформирован:" & vbCrLf & outPath & vbCrLf & vbCrLf & _
           "В базе не хранятся: объём/заводской № объекта, средства контроля, результаты " & _
           "измерений, применённые нормативные документы, ФИО руководителя работ — эти поля " & _
           "остались как #Плейсхолдеры, дозаполните их в открывшемся документе вручную.", _
           vbInformation, "Готово"

    On Error Resume Next
    Shell "explorer.exe /select,""" & outPath & """", vbNormalFocus
    Exit Sub

CleanFail:
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
