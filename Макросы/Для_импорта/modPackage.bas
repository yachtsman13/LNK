Attribute VB_Name = "modPackage"
Option Explicit

' =======================================================================
'  —борка и печать пакета документов по одной работе.
'
'  ѕор€док в пакете задан вами и не мен€етс€:
'    сводный лист-опись (если объектов в акте больше одного) -> акт ->
'    протоколы -> свидетельства лаборатории -> квалификационные
'    удостоверени€ -> свидетельства о поверке -> пустой лист дл€ сшива.
'
'  —водный лист Ч это ƒќЅј¬ ј к отдельным документам, не замена: у
'  каждого объекта и метода по-прежнему свой самосто€тельный документ,
'  а сводный лист впереди просто перечисл€ет их все одной таблицей Ч
'  чтобы не листать пакет в поисках, что в него вошло.
'
'  —квозна€ нумераци€ внутри экземпл€ра, на первой странице номера нет.
'  “ри экземпл€ра: каждый нумеруетс€ заново, поэтому печатаютс€ копии
'  одного документа, а не три склеенных подр€д.
'
'  —видетельства и удостоверени€ Ч сканы, часть в JPG/PNG/TIFF, часть в
'  PDF (сканируют разные люди по-разному, и это нормально).  артинки
'  Word вставл€ет сам Ч их страницы пр€мо в своде, в общей нумерации.
'  PDF Word вставить не может; их вместо этого объедин€ет через Acrobat
'  modScans в один файл Ч печатаетс€ и подшиваетс€ отдельно, без
'  сквозного номера (см. modScans про то, почему номер туда программно
'  не ставитс€).
' =======================================================================

Private Const wdPageBreak As Long = 7
Private Const wdStory As Long = 6
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdAlignPageNumberCenter As Long = 1
Private Const wdFieldNumPages As Long = 26
Private Const wdStatisticPages As Long = 2
Private Const wdFieldEmpty As Long = -1
Private Const wdFieldPage As Long = 33
Private Const wdReplaceOne As Long = 1
Private Const wdFindStop As Long = 0
Private Const wdAlignParagraphCenter As Long = 1


Public Sub Ќапечататьѕакет()
    Dim ws As Worksheet
    Dim r As Long
    Dim файлы As Collection
    Dim строки As Collection
    Dim акт As String
    Dim сканов As Long
    Dim wdApp As Object
    Dim свод As Object
    Dim страниц—вода As Long
    Dim страниц—канов As Long
    Dim путь—канов As String
    Dim путьPDF As String
    Dim ошибкаPDF As String
    Dim нумеровать As Boolean
    Dim всего As Long
    Dim экземпл€ров As Long
    Dim путь—вода As String

    Set ws = Ћист(Ћ_∆”–ЌјЋ)
    If ws Is Nothing Then Exit Sub
    r = “екуща€—трока(ws)
    If r = 0 Then Exit Sub

    ' ѕакет собираетс€ по ј “” целиком: акт один на работу, а объектов
    ' в ней несколько, и протоколы по всем из них идут в один пакет.
    ' ≈сли акта нет Ч печатаетс€ пакет по одной строке.
    акт = Ќомерјкта(ws, r)
    Set строки = —трокијкта(ws, r, акт)
    Set файлы = —обрать‘айлыјкта(ws, строки)
    If файлы.Count = 0 Then
        MsgBox "Ќет сформированных документов" & _
               IIf(акт = "", " по этой строке.", " по акту " & акт & ".") & vbCrLf & vbCrLf & _
               "—начала Ђ—формировать документыї Ч они по€в€тс€ в папке Ђ√отовыеї. " & _
               "≈сли в акте несколько объектов, документы нужны по каждому.", _
               vbExclamation, "ѕечать пакета"
        Exit Sub
    End If

    Application.StatusBar = "—обираю пакетЕ"
    On Error GoTo —бой
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    ' —водный лист-опись Ч только когда в акте больше одного объекта:
    ' на один объект и так есть акт/протокол, второй такой же список
    ' был бы лишним листом ни о чЄм.
    путь—вода = —обрать—водныйЋист(wdApp, ws, строки, акт)
    If путь—вода <> "" Then файлы.Add путь—вода, Before:=1

    ' —каны-картинки идут в свод как обычные страницы: свидетельства
    ' лаборатории, удостоверени€, свидетельства о поверке Ч в этом
    ' пор€дке, после протоколов. ¬кладывать их руками и держать вне
    ' нумерации больше не нужно.
    путь—канов = modScans.—обрать—каныѕакета(wdApp, строки, страниц—канов, _
                                             путьPDF, сканов, ошибкаPDF)
    If путь—канов <> "" Then файлы.Add путь—канов

    ' PDF-сканы объедин€ютс€ той же процедурой в один файл, и Acrobat
    ' сам сообщает число страниц Ч спрашивать человека не нужно, он
    ' его и не знает, пока не откроет файл. —просим только если
    ' объединение не удалось (Acrobat не найден, все файлы повреждены):
    ' тогда считать приходитс€ по старинке, как раньше.
    If путьPDF = "" And ошибкаPDF <> "" Then
        сканов = —просить„исло—канов()
        If сканов < 0 Then
            wdApp.Quit
            Exit Sub
        End If
    End If

    Set свод = —клеитьƒокументы(wdApp, файлы)

    ƒобавитьѕустойЋист свод
    страниц—вода = свод.ComputeStatistics(wdStatisticPages)
    нумеровать = ѕоЋистам()
    ѕроставитьЌомера свод, нумеровать
    If нумеровать Then
        ' ƒвусторонн€€ печать: лист Ч это две страницы, и последн€€
        ' нечЄтна€ страница всЄ равно занимает целый лист.
        всего = (страниц—вода + 1) \ 2 + сканов
    Else
        всего = страниц—вода + сканов
    End If

    Application.StatusBar = False
    MsgBox "ѕакет собран." & vbCrLf & vbCrLf & _
           IIf(акт = "", "Ѕез акта, одна строка журнала." & vbCrLf, _
               акт & ", объектов в нЄм: " & строки.Count & vbCrLf) & _
           "ƒокументов Word: " & файлы.Count & vbCrLf & _
           "—траниц в них (с пустым листом): " & страниц—вода & vbCrLf & _
           IIf(страниц—канов = 0, "", "»з них страниц сканов-картинок: " & страниц—канов & _
               " Ч они уже в своде и в общей нумерации" & vbCrLf) & _
           IIf(путьPDF = "", "", "PDF-сканы объединены Acrobat в " & сканов & _
               " стр.: " & путьPDF & vbCrLf) & _
           IIf(путьPDF = "" And сканов > 0, "—траниц PDF, вкладываемых вручную: " & _
               сканов & vbCrLf, "") & _
           "Ќумераци€: " & IIf(нумеровать, "по листам, номер на лицевой стороне", _
                               "по страницам, сквозна€") & vbCrLf & _
           "¬сего листов в экземпл€ре: " & всего, _
           vbInformation, "ѕечать пакета"

    экземпл€ров = —просить„ислоЁкземпл€ров()
    If экземпл€ров < 0 Then
        свод.Close SaveChanges:=False
        wdApp.Quit
        Exit Sub
    End If

    —охранить—вод свод, ws, r
    If экземпл€ров > 0 Then
        ' Copies:=N печатает N отдельных экземпл€ров, и нумераци€ в
        ' каждом начинаетс€ заново Ч именно это и нужно.
        свод.PrintOut Copies:=экземпл€ров, Collate:=True
        ЌапечататьЅирку wdApp, всего, ws, r, экземпл€ров
    End If

    wdApp.Visible = True
    MsgBox "√отово." & vbCrLf & vbCrLf & _
           IIf(страниц—канов > 0, "—каны-картинки вошли в пакет и пронумерованы вместе " & _
               "со всем остальным." & vbCrLf & vbCrLf, "") & _
           IIf(путьPDF <> "", "PDF-сканы объединены в один файл Ч откройте его в Acrobat, " & _
               "напечатайте (" & IIf(экземпл€ров > 0, экземпл€ров & " экз., двусторонне", _
                                     "столько экземпл€ров, сколько будете печатать пакет") & _
               ") и вложите между протоколами и пустым листом: " & путьPDF & _
               IIf(ошибкаPDF <> "", " (" & ошибкаPDF & ")", "") & vbCrLf & vbCrLf, "") & _
           IIf(путьPDF = "" And ошибкаPDF <> "", "PDF-сканы найдены, но Acrobat их не " & _
               "объединил: " & ошибкаPDF & vbCrLf & "¬ложите их по одному, как раньше." & _
               vbCrLf & vbCrLf, "") & _
           IIf(страниц—канов = 0 And путьPDF = "" And ошибкаPDF = "", _
               "—канов не нашлось Ч проверьте папку сканов в настройках.", ""), _
           vbInformation, "ѕечать пакета"
    Exit Sub

—бой:
    Application.StatusBar = False
    On Error Resume Next
    If Not свод Is Nothing Then свод.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    MsgBox "Ќе удалось собрать пакет." & vbCrLf & "ќшибка: " & Err.Description, _
           vbCritical, "ѕечать пакета"
End Sub


' -- —бор файлов -------------------------------------------------------

' ¬се строки одного акта. Ѕез номера акта Ч только текуща€ строка.
Public Function —трокијкта(ws As Worksheet, r As Long, акт As String) As Collection
    Dim список As Collection
    Dim i As Long, last As Long

    Set список = New Collection
    Set —трокијкта = список
    If акт = "" Then
        список.Add r
        Exit Function
    End If

    last = ѕоследн€€—трока(ws, ColReq(ws, "ƒата контрол€"))
    For i = ѕ≈–¬јя_—“–ќ ј To last
        If StrComp(Ќомерјкта(ws, i), акт, vbTextCompare) = 0 Then список.Add i
    Next i
End Function


' ѕор€док в пакете: сначала акт ¬»  по всем объектам, затем протоколы
' остальных методов Ч как задано в вашем пор€дке сшива.
Private Function —обрать‘айлыјкта(ws As Worksheet, строки As Collection) As Collection
    Dim список As Collection
    Dim методы As Variant
    Dim i As Long, k As Long
    Dim метод As String
    Dim путь As String

    Set список = New Collection
    Set —обрать‘айлыјкта = список

    методы = Split(ћ≈“ќƒџ, ",")
    For i = LBound(методы) To UBound(методы)
        метод = CStr(методы(i))
        For k = 1 To строки.Count
            If ћетодѕрименЄн(ws, CLng(строки(k)), метод) Then
                путь = Ќайтиƒокумент(ws, CLng(строки(k)), метод)
                If путь <> "" Then список.Add путь
            End If
        Next k
    Next i
End Function


Private Function Ќайтиƒокумент(ws As Worksheet, r As Long, метод As String) As String
    Dim папка As String
    Dim маска As String
    Dim файл As String
    Dim дата As Variant
    Dim год As Long

    дата = ws.Cells(r, ColReq(ws, "ƒата контрол€")).Value
    год = IIf(IsDate(дата), Year(CDate(дата)), Year(Date))
    папка = ѕапка√отовых() & CStr(год) & Application.PathSeparator & метод & Application.PathSeparator
    If Not ѕапка—уществует(папка) Then Exit Function

    маска = ќчиститьƒл€»мени‘айла(–егЌомер(ws, r, метод)) & "*.docx"
    файл = ѕервый‘айл(папка & маска)
    If файл <> "" Then Ќайтиƒокумент = папка & файл
End Function


' -- —водный лист-опись протоколов акта ----------------------------------
'
' “олько когда в акте больше одного объекта: печатаетс€ впереди пакета,
' одним текстом перечисл€ет все документы акта, сгруппированные по
' методу Ч как в вашем образце (Ђјкты ¬»  єє Еї, Ђѕротоколы ”«“ єє Еї).
' Ёто ƒќЅј¬ ј к отдельным документам по каждому объекту и методу, не
' замена: они остаютс€ в пакете как есть.
'
' „асть текста (номер и дата нар€да-заказа) система по данным не
' отслеживает Ч как и на бирке, оставлены пустые места под запись от
' руки. ѕомен€ть формулировку целиком можно настройкой Ђ“екст сводного
' актаї, не трога€ код.
Private Function —обрать—водныйЋист(wdApp As Object, ws As Worksheet, строки As Collection, _
                                     акт As String) As String
    Dim методы As Variant
    Dim объекты As String, документы As String
    Dim текст As String
    Dim номерјкта As String
    Dim док As Object
    Dim исполнители As Collection
    Dim подпись“абл As Object
    Dim должность As String, кратко As String
    Dim k As Long
    Dim путь As String

    If строки.Count <= 1 Then Exit Function

    методы = Split(ћ≈“ќƒџ, ",")
    документы = —обратьƒокументыјкта(ws, строки, методы)
    If документы = "" Then Exit Function     ' ни одного зарегистрированного номера

    объекты = —обратьќбъектыјкта(ws, строки)

    номерјкта = акт
    If InStr(1, номерјкта, "ј “ ", vbTextCompare) = 1 Then номерјкта = Mid$(номерјкта, 5)

    текст = Ќастройка("“екст сводного акта")
    If текст = "" Then
        текст = "¬ соответствии с нар€дом-заказом (за€вкой) є ______ от " & _
                "Ђ____ї ________________ 20____ г. выполнены работы по " & _
                "неразрушающему контролю на объектах: #ќЅЏ≈ “џ." & vbCr & vbCr & _
                "–езультаты работ отражены в следующих документах." & vbCr & vbCr & _
                "#ƒќ ”ћ≈Ќ“џ"
    End If
    текст = Replace(текст, "#ј “", номерјкта)
    текст = Replace(текст, "#ќЅЏ≈ “џ", объекты)
    текст = Replace(текст, "#ƒќ ”ћ≈Ќ“џ", документы)

    Set док = wdApp.Documents.Add
    With док.Paragraphs(1).Range
        .Text = "ј “ є " & номерјкта
        .Font.Bold = True
        .Font.Size = 14
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
    док.Content.InsertParagraphAfter
    With док.Paragraphs(2).Range
        .Text = "о проведении работ"
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
    док.Content.InsertParagraphAfter
    док.Content.InsertParagraphAfter

    With док.Paragraphs(док.Paragraphs.Count).Range
        .InsertAfter текст
    End With
    док.Content.InsertParagraphAfter
    док.Content.InsertParagraphAfter

    ' ѕодписи Ч те, кто по журналу исполн€л методы этого акта; должность
    ' подт€гиваетс€ из реестра специалистов, а не вписываетс€ вручную.
    Set исполнители = »сполнителијкта(ws, строки, методы)
    If исполнители.Count > 0 Then
        Set подпись“абл = док.Tables.Add(Range:=док.Paragraphs(док.Paragraphs.Count).Range, _
                                         NumRows:=исполнители.Count, NumColumns:=2)
        подпись“абл.Borders.Enable = False
        For k = 1 To исполнители.Count
            должность = ""
            кратко = ƒанные—пециалиста(CStr(исполнители(k)), должность)
            подпись“абл.Cell(k, 1).Range.Text = должность
            подпись“абл.Cell(k, 2).Range.Text = кратко
        Next k
    End If

    путь = ѕапка√отовых() & "—водный_" & ќчиститьƒл€»мени‘айла(номерјкта) & "_" & _
           Format$(Now, "yyyymmdd_hhnnss") & ".docx"
    док.SaveAs2 fileName:=путь
    док.Close SaveChanges:=False
    —обрать—водныйЋист = путь
End Function


Private Function ќбъект—–еквизитами(ws As Worksheet, r As Long) As String
    Dim им€ As String, зав As String, рег As String
    им€ = Trim$(CStr(ws.Cells(r, ColReq(ws, "Ќаименование объекта в документе")).Value))
    зав = Trim$(CStr(ws.Cells(r, ColReq(ws, "«аводской є объекта")).Value))
    рег = Trim$(CStr(ws.Cells(r, ColReq(ws, "–ег. / уч. є объекта")).Value))
    If зав <> "" Then им€ = им€ & ", зав. є " & зав
    If рег <> "" Then им€ = им€ & ", рег. є " & рег
    ќбъект—–еквизитами = им€
End Function


' ќбъекты через Ђ; ї, последний Ч через Ђ и ї, как в образце заказчика.
Private Function —обратьќбъектыјкта(ws As Worksheet, строки As Collection) As String
    Dim i As Long
    Dim s As String
    For i = 1 To строки.Count
        If i > 1 Then s = s & IIf(i = строки.Count, " и ", "; ")
        s = s & ќбъект—–еквизитами(ws, CLng(строки(i)))
    Next i
    —обратьќбъектыјкта = s
End Function


' Ќазвание типа документа во множественном числе Ч как в шапке ваших
' шаблонов, только во множественном числе и без номера. ƒл€ ћ  шаблона
' документа в системе ещЄ нет (известный пробел) Ч название подставлено
' по аналогии с остальными методами и не подтверждено.
Private Function “ипƒокументаћн(метод As String) As String
    Select Case UCase$(метод)
        Case "¬» ": “ипƒокументаћн = "јкты визуального и (или) измерительного контрол€"
        Case "”«ƒ": “ипƒокументаћн = "«аключени€ по результатам ультразвукового контрол€"
        Case "”«“": “ипƒокументаћн = "ѕротоколы ультразвуковой толщинометрии"
        Case "јЁ":  “ипƒокументаћн = "«аключени€ по результатам акустико-эмиссионного контрол€"
        Case "“¬":  “ипƒокументаћн = "ѕротоколы измерени€ твЄрдости металла"
        Case "“ ":  “ипƒокументаћн = "ѕротоколы теплового контрол€"
        Case "¬ƒ":  “ипƒокументаћн = "ѕротоколы вибродиагностического контрол€"
        Case "ћ ":  “ипƒокументаћн = "ѕротоколы магнитопорошкового контрол€"
        Case Else:  “ипƒокументаћн = "ƒокументы по методу " & метод
    End Select
End Function


Private Function —обратьƒокументыјкта(ws As Worksheet, строки As Collection, методы As Variant) As String
    Dim i As Long, m As Long
    Dim метод As String, номер As String
    Dim номера As Collection
    Dim блоки As New Collection

    For m = LBound(методы) To UBound(методы)
        метод = CStr(методы(m))
        Set номера = New Collection
        For i = 1 To строки.Count
            If ћетодѕрименЄн(ws, CLng(строки(i)), метод) Then
                номер = –егЌомер(ws, CLng(строки(i)), метод)
                If номер <> "" Then номера.Add номер
            End If
        Next i
        If номера.Count > 0 Then
            блоки.Add “ипƒокументаћн(метод) & " єє " & —оединить«ап€тыми(номера) & "."
        End If
    Next m

    —обратьƒокументыјкта = —оединитьѕереносами(блоки)
End Function


Private Function —оединить«ап€тыми(список As Collection) As String
    Dim i As Long, s As String
    For i = 1 To список.Count
        If i > 1 Then s = s & ", "
        s = s & CStr(список(i))
    Next i
    —оединить«ап€тыми = s
End Function


Private Function —оединитьѕереносами(список As Collection) As String
    Dim i As Long, s As String
    For i = 1 To список.Count
        If i > 1 Then s = s & vbCr & vbCr
        s = s & CStr(список(i))
    Next i
    —оединитьѕереносами = s
End Function


'  то исполн€л методы этого акта Ч без повторов, в пор€дке по€влени€ в
' журнале. ƒолжность подт€гиваетс€ по ‘»ќ из реестра специалистов.
Private Function »сполнителијкта(ws As Worksheet, строки As Collection, методы As Variant) As Collection
    Dim i As Long, m As Long, k As Long
    Dim метод As String, фио As String
    Dim найден As Boolean
    Dim список As New Collection

    For i = 1 To строки.Count
        For m = LBound(методы) To UBound(методы)
            метод = CStr(методы(m))
            If ћетодѕрименЄн(ws, CLng(строки(i)), метод) Then
                фио = Trim$(CStr(ws.Cells(CLng(строки(i)), ColReq(ws, метод & " исполнитель")).Value))
                If фио <> "" Then
                    найден = False
                    For k = 1 To список.Count
                        If StrComp(CStr(список(k)), фио, vbTextCompare) = 0 Then
                            найден = True
                            Exit For
                        End If
                    Next k
                    If Not найден Then список.Add фио
                End If
            End If
        Next m
    Next i
    Set »сполнителијкта = список
End Function


' ‘»ќ в короткой форме (Ђј.¬. »вановї) и должность из реестра
' специалистов Ч по полному или короткому ‘»ќ, как записано в журнале.
' Ќе нашли Ч возвращаетс€ то же ‘»ќ, что и было, должность пуста€.
Private Function ƒанные—пециалиста(фио As String, ByRef должность As String) As String
    Dim sp As Worksheet
    Dim i As Long, last As Long
    Dim полностью As String, кратко As String

    ƒанные—пециалиста = фио
    Set sp = Ћист(Ћ_—ѕ≈÷»јЋ»—“џ)
    If sp Is Nothing Then Exit Function

    last = ѕоследн€€—трока(sp, ColReq(sp, "ID"))
    For i = ѕ≈–¬јя_—“–ќ ј To last
        полностью = Trim$(CStr(sp.Cells(i, ColReq(sp, "‘»ќ полностью")).Value))
        кратко = Trim$(CStr(sp.Cells(i, ColReq(sp, "‘»ќ в документах")).Value))
        If StrComp(полностью, фио, vbTextCompare) = 0 Or StrComp(кратко, фио, vbTextCompare) = 0 Then
            должность = Trim$(CStr(sp.Cells(i, ColReq(sp, "ƒолжность")).Value))
            If кратко <> "" Then ƒанные—пециалиста = кратко
            Exit Function
        End If
    Next i
End Function


' -- —борка документа --------------------------------------------------

Private Function —клеитьƒокументы(wdApp As Object, файлы As Collection) As Object
    Dim свод As Object
    Dim i As Long

    Set свод = wdApp.Documents.Add
    For i = 1 To файлы.Count
        If i > 1 Then
            свод.Content.InsertParagraphAfter
            свод.Paragraphs(свод.Paragraphs.Count).Range.InsertBreak Type:=wdPageBreak
        End If
        свод.Content.InsertParagraphAfter
        свод.Paragraphs(свод.Paragraphs.Count).Range.InsertFile fileName:=CStr(файлы(i))
    Next i
    Set —клеитьƒокументы = свод
End Function


' ѕоследний лист пакета Ч абсолютно пустой, кроме номера страницы:
' он нужен, чтобы прошить пакет и наклеить бирку.
Private Sub ƒобавитьѕустойЋист(свод As Object)
    свод.Content.InsertParagraphAfter
    свод.Paragraphs(свод.Paragraphs.Count).Range.InsertBreak Type:=wdPageBreak
    свод.Content.InsertParagraphAfter
End Sub


Private Function ѕоЋистам() As Boolean
    Dim s As String
    s = Ќастройка("Ќумеровать")
    ѕоЋистам = (s = "" Or StrComp(Left$(s, 4), "Ћист", vbTextCompare) = 0)
End Function


' Ќомер по центру нижнего колонтитула, на первой странице номера нет.
'
' ѕо листам Ч это нумераци€ подшивки, а не документа: при двусторонней
' печати лист бумаги несЄт две страницы, а номер у него один. ѕоэтому
' номер ставитс€ только на нечЄтных (лицевых) страницах и считает не
' страницы, а листы: 1, 3, 5-€ страницы получают 1, 2, 3.
'
' —читает это поле {=INT((PAGE+1)/2)} Ч вложенное поле внутри формулы.
' —обрать его из VBA одной строкой нельз€, поэтому собираетс€ в два
' приЄма и результат провер€етс€: если Word его не пон€л, ставитс€
' обычна€ нумераци€ страниц и об этом говоритс€ вслух Ч молча
' напечатать пакет не с той нумерацией хуже, чем сказать.
Private Sub ѕроставитьЌомера(свод As Object, режимЋистов As Boolean)
    Dim s As Object
    Dim получилось As Boolean

    получилось = режимЋистов
    If режимЋистов Then
        For Each s In свод.Sections
            s.PageSetup.DifferentFirstPageHeaderFooter = True
            s.PageSetup.OddAndEvenPagesHeaderFooter = True
            ' Primary при включЄнных чЄтных/нечЄтных Ч это нечЄтные, то
            ' есть лицевые стороны листов. „Єтный колонтитул пуст.
            If Not ЌомерЋиста(s.Footers(wdHeaderFooterPrimary)) Then
                получилось = False
                Exit For
            End If
        Next s
    End If

    If получилось Then Exit Sub

    For Each s In свод.Sections
        s.PageSetup.DifferentFirstPageHeaderFooter = True
        s.PageSetup.OddAndEvenPagesHeaderFooter = False
        With s.Footers(wdHeaderFooterPrimary)
            .Range.Text = ""
            .PageNumbers.RestartNumberingAtSection = False
            .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=False
        End With
    Next s

    If режимЋистов Then
        MsgBox "Ќумераци€ по листам не собралась Ч Word не прин€л поле " & _
               "{=INT((PAGE+1)/2)}." & vbCrLf & vbCrLf & _
               "ѕакет пронумерован по страницам: печатать так можно, но это " & _
               "не то, что просили. —кажите Ч разберЄмс€, почему Word отказалс€.", _
               vbExclamation, "ѕечать пакета"
    End If
End Sub


' ¬кладывает поле PAGE внутрь формулы. True Ч получилось.
Private Function ЌомерЋиста(колонтитул As Object) As Boolean
    Dim rng As Object
    Dim поле As Object
    Dim код As Object
    Dim результат As String

    On Error GoTo Ќет
    Set rng = колонтитул.Range
    rng.Text = ""
    Set rng = колонтитул.Range
    Set поле = rng.Fields.Add(rng, wdFieldEmpty, "=INT((—“–јЌ»÷ј+1)/2)", False)

    Set код = поле.Code
    With код.Find
        .Text = "—“–јЌ»÷ј"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceOne
    End With
    ' ѕосле замены диапазон схлопнулс€ там, где сто€ло слово, Ч туда и
    ' вкладываетс€ насто€щее поле PAGE.
    код.Fields.Add код, wdFieldPage, , False
    поле.Update

    колонтитул.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    результат = Trim$(поле.Result.Text)
    ЌомерЋиста = (результат <> "" And IsNumeric(результат))
    Exit Function
Ќет:
    колонтитул.Range.Text = ""
End Function


Private Sub —охранить—вод(свод As Object, ws As Worksheet, r As Long)
    Dim папка As String
    Dim им€ As String
    Dim дата As Variant
    Dim год As Long

    On Error Resume Next
    дата = ws.Cells(r, ColReq(ws, "ƒата контрол€")).Value
    год = IIf(IsDate(дата), Year(CDate(дата)), Year(Date))
    папка = ѕапка√отовых() & CStr(год) & Application.PathSeparator & "ѕакеты" & Application.PathSeparator
    MkPath папка
    им€ = Ќомерјкта(ws, r)
    If им€ = "" Then
        им€ = CStr(ws.Cells(r, ColReq(ws, "ID записи")).Value) & " " & _
              Left$(CStr(ws.Cells(r, ColReq(ws, "Ќаименование объекта в документе")).Value), 50)
    End If
    им€ = ќчиститьƒл€»мени‘айла(им€)
    свод.SaveAs2 fileName:=папка & им€ & ".docx"
End Sub


' -- Ѕирка -------------------------------------------------------------

' „исло листов на бирке пишетс€ цифрами и прописью. ѕрописью его чаще
' всего и пишут с ошибкой, поэтому склон€етс€ автоматически.
Private Sub ЌапечататьЅирку(wdApp As Object, листов As Long, ws As Worksheet, r As Long, _
                            экземпл€ров As Long)
    Dim док As Object
    Dim текст As String
    Dim шаблон As String

    On Error Resume Next
    шаблон = Ќастройка("“екст бирки")
    If шаблон = "" Then
        шаблон = "ѕрошнуровано, пронумеровано и скреплено печатью" & vbCrLf & _
                 "#„»—Ћќ (#ѕ–ќѕ»—№ё) #Ћ»—“ќ¬" & vbCrLf & vbCrLf & _
                 "ƒолжность ______________________" & vbCrLf & vbCrLf & _
                 "ѕодпись, ‘»ќ ___________________" & vbCrLf & vbCrLf & _
                 "Ђ____ї ________________ 20____ г."
    End If

    текст = Replace(шаблон, "#„»—Ћќ", CStr(листов))
    текст = Replace(текст, "#ѕ–ќѕ»—№ё", „ислоѕрописью(листов))
    текст = Replace(текст, "#Ћ»—“ќ¬", —клонениеЋистов(листов))
    текст = Replace(текст, "#ќЅЏ≈ “", CStr(ws.Cells(r, ColReq(ws, "Ќаименование объекта в документе")).Value))

    Set док = wdApp.Documents.Add
    док.Content.Text = текст
    док.Content.ParagraphFormat.SpaceAfter = 6
    док.PrintOut Copies:=экземпл€ров
End Sub


Public Function „ислоѕрописью(n As Long) As String
    Dim сотни As Variant
    Dim дес€тки As Variant
    Dim до20 As Variant
    Dim с As Long, д As Long, е As Long
    Dim s As String

    If n <= 0 Then Exit Function
    If n > 999 Then
        „ислоѕрописью = CStr(n)          ' пакетов такой толщины не бывает
        Exit Function
    End If

    сотни = Array("", "сто", "двести", "триста", "четыреста", "п€тьсот", _
                  "шестьсот", "семьсот", "восемьсот", "дев€тьсот")
    дес€тки = Array("", "", "двадцать", "тридцать", "сорок", "п€тьдес€т", _
                    "шестьдес€т", "семьдес€т", "восемьдес€т", "дев€носто")
    до20 = Array("", "один", "два", "три", "четыре", "п€ть", "шесть", "семь", _
                 "восемь", "дев€ть", "дес€ть", "одиннадцать", "двенадцать", _
                 "тринадцать", "четырнадцать", "п€тнадцать", "шестнадцать", _
                 "семнадцать", "восемнадцать", "дев€тнадцать")

    с = n \ 100
    д = (n - с * 100) \ 10
    е = n - с * 100 - д * 10

    If с > 0 Then s = сотни(с)
    If д = 1 Then
        s = Trim$(s & " " & до20(10 + е))
    Else
        If д > 1 Then s = Trim$(s & " " & дес€тки(д))
        If е > 0 Then s = Trim$(s & " " & до20(е))
    End If
    „ислоѕрописью = Trim$(s)
End Function


Public Function —клонениеЋистов(n As Long) As String
    Dim последн€€ As Long
    Dim две As Long
    последн€€ = n Mod 10
    две = n Mod 100
    If две >= 11 And две <= 14 Then
        —клонениеЋистов = "листов"
    ElseIf последн€€ = 1 Then
        —клонениеЋистов = "лист"
    ElseIf последн€€ >= 2 And последн€€ <= 4 Then
        —клонениеЋистов = "листа"
    Else
        —клонениеЋистов = "листов"
    End If
End Function


Private Function —просить„исло—канов() As Long
    Dim s As String
    s = InputBox("—колько листов вкладываетс€ в пакет ¬–”„Ќ”ё?" & vbCrLf & vbCrLf & _
                 "—каны-картинки (JPG, PNG, TIFF) программа вставл€ет в пакет сама Ч " & _
                 "их здесь считать не надо. –уками остаютс€ только сканы в PDF: " & _
                 "объединить PDF без Acrobat нельз€." & vbCrLf & vbCrLf & _
                 "Ёто число нужно, чтобы на бирке сошлось количество листов.", _
                 "ѕечать пакета", "0")
    If StrPtr(s) = 0 Then                ' нажали Ђќтменаї
        —просить„исло—канов = -1
        Exit Function
    End If
    If Trim$(s) = "" Then Exit Function
    If Not IsNumeric(s) Then
        MsgBox "Ќужно число.", vbExclamation, "ѕечать пакета"
        —просить„исло—канов = -1
        Exit Function
    End If
    —просить„исло—канов = CLng(s)
End Function


' —колько экземпл€ров печатать Ч раньше было жЄстко три, без выбора.
' 0 Ч собрать и сохранить файл, но не печатать сейчас (например, дл€
' проверки перед печатью или если принтер сейчас недоступен). ќтмена Ч
' не сохран€ть вовсе, свод закрываетс€ без следа.
Private Function —просить„ислоЁкземпл€ров() As Long
    Dim s As String
    s = InputBox("—колько экземпл€ров печатать?" & vbCrLf & vbCrLf & _
                 "0 Ч собрать и сохранить файл, но не печатать сейчас.", _
                 "ѕечать пакета", "3")
    If StrPtr(s) = 0 Then                ' нажали Ђќтменаї
        —просить„ислоЁкземпл€ров = -1
        Exit Function
    End If
    s = Trim$(s)
    If Not IsNumeric(s) Or CLng(Val(s)) < 0 Then
        MsgBox "Ќужно неотрицательное число.", vbExclamation, "ѕечать пакета"
        —просить„ислоЁкземпл€ров = -1
        Exit Function
    End If
    —просить„ислоЁкземпл€ров = CLng(s)
End Function
