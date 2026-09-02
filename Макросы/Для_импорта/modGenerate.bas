Attribute VB_Name = "modGenerate"
Option Explicit

' =======================================================================
'  Формирование документов Word из шаблонов.
'
'  В шаблонах поля помечены как #ИмяПоля — так же, как в ваших рабочих
'  шаблонах. Что система знает, она подставляет; чего не знает —
'  оставляет плейсхолдером и честно перечисляет в конце, чтобы
'  незаполненное не уехало в документ незамеченным.
' =======================================================================

Private Const wdReplaceAll As Long = 2
Private Const wdFindContinue As Long = 1


' Формирует документы по всем методам текущей строки журнала.
Public Sub СформироватьДокументы()
    Dim ws As Worksheet
    Dim r As Long
    Dim методы As Variant
    Dim i As Long
    Dim метод As String
    Dim сделано As String
    Dim путь As String
    Dim предупреждения As String
    Dim безШаблона As String

    Set ws = Лист(Л_ЖУРНАЛ)
    If ws Is Nothing Then Exit Sub
    r = ТекущаяСтрока(ws)
    If r = 0 Then Exit Sub

    предупреждения = ПроверитьСтрокуЖурнала(r)
    If предупреждения <> "" Then
        If MsgBox("Перед формированием стоит посмотреть:" & vbCrLf & предупреждения & vbCrLf & vbCrLf & _
                  "Формировать всё равно?", vbYesNo + vbExclamation, "Проверка") = vbNo Then Exit Sub
    End If

    методы = Split(МЕТОДЫ, ",")
    For i = LBound(методы) To UBound(методы)
        метод = методы(i)
        If МетодПрименён(ws, r, метод) Then
            If ФайлШаблона(метод) = "" Then
                безШаблона = безШаблона & IIf(безШаблона = "", "", ", ") & метод
            Else
                путь = СформироватьОдин(ws, r, метод)
                If путь <> "" Then сделано = сделано & vbCrLf & "  " & путь
            End If
        End If
    Next i

    If сделано = "" Then
        MsgBox "Ни одного документа не сформировано." & vbCrLf & _
               IIf(безШаблона = "", "В строке не отмечен ни один метод контроля.", _
                   "Для методов " & безШаблона & " шаблонов пока нет — эти документы " & _
                   "оформляются вручную."), vbExclamation, "Формирование документов"
    Else
        ws.Cells(r, ColReq(ws, "Документ сформирован")).Value = Date
        MsgBox "Сформировано:" & сделано & _
               IIf(безШаблона = "", "", vbCrLf & vbCrLf & "Без шаблона остались методы: " & _
                   безШаблона & " — эти документы оформляются вручную.") & vbCrLf & vbCrLf & _
               "Незаполненные поля остались в виде #Плейсхолдеров — допишите их в документе.", _
               vbInformation, "Формирование документов"
    End If
End Sub


Private Function СформироватьОдин(ws As Worksheet, r As Long, метод As String) As String
    Dim шаблон As String
    Dim путьШаблона As String
    Dim папка As String
    Dim файл As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim дата As Variant
    Dim год As Long

    шаблон = ФайлШаблона(метод)
    If шаблон = "" Then Exit Function
    путьШаблона = ПапкаШаблонов() & шаблон
    If Not ФайлСуществует(путьШаблона) Then
        MsgBox "Не найден шаблон для метода " & метод & ":" & vbCrLf & путьШаблона, _
               vbExclamation, "Формирование документов"
        Exit Function
    End If

    дата = ws.Cells(r, ColReq(ws, "Дата контроля")).Value
    год = IIf(IsDate(дата), Year(CDate(дата)), Year(Date))

    On Error GoTo Сбой
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(fileName:=путьШаблона, ReadOnly:=False, AddToRecentFiles:=False)

    Подставить wdDoc, "Номер", РегНомер(ws, r, метод)
    Подставить wdDoc, "Год", CStr(год)
    Подставить wdDoc, "Вид_контроля", метод
    Подставить wdDoc, "Дата_и_место_проведения_контроля", ДатаИМесто(ws, r)
    Подставить wdDoc, "Объект_контроля", ОписаниеОбъекта(ws, r)
    Подставить wdDoc, "Наименование_объекта", ОписаниеОбъекта(ws, r)
    Подставить wdDoc, "Заказывающее_подразделение", _
               CStr(ws.Cells(r, ColReq(ws, "Заказывающее подразделение")).Value)
    Подставить wdDoc, "Уровень_квалификации_№_удостоверения_ФИО", ПодписьИсполнителя(ws, r, метод)
    Подставить wdDoc, "Уровень_№_ФИО", ПодписьИсполнителя(ws, r, метод)
    Подставить wdDoc, "Наименование_ЛНК", Настройка("Полное наименование ЛНК")

    папка = ПапкаГотовых() & CStr(год) & Application.PathSeparator & метод & Application.PathSeparator
    MkPath папка
    файл = папка & ОчиститьДляИмениФайла(РегНомер(ws, r, метод)) & " " & _
           ОчиститьДляИмениФайла(Left$(ОписаниеОбъекта(ws, r), 60)) & ".docx"

    wdDoc.SaveAs2 fileName:=файл
    wdDoc.Close SaveChanges:=False
    wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing

    СформироватьОдин = файл
    Exit Function

Сбой:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    MsgBox "Не удалось сформировать документ по методу " & метод & "." & vbCrLf & _
           "Ошибка: " & Err.Description, vbCritical, "Формирование документов"
End Function


Private Sub Подставить(doc As Object, поле As String, значение As String)
    Dim rng As Object
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Text = "#" & поле
        .Replacement.ClearFormatting
        .Replacement.Text = значение
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub


Public Function ФайлШаблона(метод As String) As String
    Select Case UCase$(Trim$(метод))
        Case "ВИК": ФайлШаблона = "Акт_ВИК.docx"
        Case "УЗД": ФайлШаблона = "Заключение_УЗД.docx"
        Case "УЗТ": ФайлШаблона = "Протокол_УЗТ.docx"
        Case "АЭ": ФайлШаблона = "Протокол_АЭ_полный.docx"
        Case "ТВ": ФайлШаблона = "Протокол_ТВ.docx"
        Case "ТК": ФайлШаблона = "Протокол_ТК.docx"
        Case "ВД": ФайлШаблона = "Протокол_ВД.docx"
        Case Else: ФайлШаблона = ""      ' МК — шаблона пока нет
    End Select
End Function


Private Function ОписаниеОбъекта(ws As Worksheet, r As Long) As String
    Dim имя As String
    Dim зав As String
    имя = Trim$(CStr(ws.Cells(r, ColReq(ws, "Наименование объекта в документе")).Value))
    зав = Trim$(CStr(ws.Cells(r, ColReq(ws, "Заводской № объекта")).Value))
    If зав <> "" Then имя = имя & ", зав. № " & зав
    ОписаниеОбъекта = имя
End Function


Private Function ДатаИМесто(ws As Worksheet, r As Long) As String
    Dim s As String
    Dim место As String
    If IsDate(ws.Cells(r, ColReq(ws, "Дата контроля")).Value) Then
        s = Format$(ws.Cells(r, ColReq(ws, "Дата контроля")).Value, "dd.mm.yyyy")
    End If
    место = Trim$(CStr(ws.Cells(r, ColReq(ws, "Заказывающее подразделение")).Value))
    If место <> "" Then
        If s <> "" Then s = s & ", "
        s = s & место
    End If
    ДатаИМесто = s
End Function


' Подпись исполнителя: уровень и номер удостоверения берутся из реестра
' аттестаций — переписывать их в каждый документ руками незачем.
Private Function ПодписьИсполнителя(ws As Worksheet, r As Long, метод As String) As String
    Dim фио As String
    Dim att As Worksheet
    Dim i As Long, last As Long
    Dim уровень As String
    Dim удост As String

    фио = Trim$(CStr(ws.Cells(r, ColReq(ws, метод & " исполнитель")).Value))
    If фио = "" Then Exit Function

    Set att = Лист(Л_АТТЕСТАЦИИ)
    If att Is Nothing Then
        ПодписьИсполнителя = фио
        Exit Function
    End If
    last = ПоследняяСтрока(att, ColReq(att, "ID"))
    For i = ПЕРВАЯ_СТРОКА To last
        If StrComp(Trim$(CStr(att.Cells(i, ColReq(att, "ФИО")).Value)), фио, vbTextCompare) = 0 And _
           StrComp(Trim$(CStr(att.Cells(i, ColReq(att, "Метод контроля")).Value)), метод, vbTextCompare) = 0 Then
            уровень = Trim$(CStr(att.Cells(i, ColReq(att, "Уровень")).Value))
            удост = Trim$(CStr(att.Cells(i, ColReq(att, "№ квалификационного удостоверения")).Value))
            Exit For
        End If
    Next i

    ПодписьИсполнителя = Trim$(IIf(уровень = "", "", уровень & " уровень, ") & _
                               IIf(удост = "", "", "уд. № " & удост & ", ") & фио)
End Function
