Attribute VB_Name = "modUpdate"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════
'  Обновление версии программы.
'
'  Заменяется только ЛНК_Программа.xlsm; ЛНК_Данные.xlsx не трогается
'  никогда. Иначе любое обновление затирало бы рабочие данные — это самая
'  частая причина, по которой такие системы перестают обновлять.
'
'  Целостность файла проверяется контрольной суммой SHA-256. Считает её
'  штатный certutil: он есть в Windows и прав администратора не требует.
'  Проверка ловит и оборвавшееся копирование по нестабильной сети, и
'  подменённый файл.
' ═══════════════════════════════════════════════════════════════════════

Private Const ФАЙЛ_ВЕРСИИ As String = "версия.txt"
Private Const ФАЙЛ_ПРОГРАММЫ As String = "ЛНК_Программа.xlsm"


' ── Проверка при запуске ──────────────────────────────────────────────

Public Sub ПроверитьОбновление(Optional молча As Boolean = True)
    Dim папка As String
    Dim номерВСети As String
    Dim мойНомер As String

    On Error Resume Next
    папка = ПапкаОбмена() & "Версия" & Application.PathSeparator
    If Not ФайлСуществует(папка & ФАЙЛ_ВЕРСИИ) Then
        If Not молча Then
            MsgBox "В общей папке нет файла версии — обновлений не публиковали.", _
                   vbInformation, "Обновление"
        End If
        Exit Sub
    End If

    номерВСети = ПолеВерсии(папка & ФАЙЛ_ВЕРСИИ, 1)
    мойНомер = Настройка("Версия программы")
    If мойНомер = "" Then мойНомер = "0"

    If СравнитьВерсии(номерВСети, мойНомер) <= 0 Then
        If Not молча Then
            MsgBox "У вас последняя версия: " & мойНомер & ".", vbInformation, "Обновление"
        End If
        Exit Sub
    End If

    If MsgBox("В общей папке лежит версия " & номерВСети & ", у вас " & мойНомер & "." & vbCrLf & _
              "Выпустил: " & ПолеВерсии(папка & ФАЙЛ_ВЕРСИИ, 3) & _
              ", " & ПолеВерсии(папка & ФАЙЛ_ВЕРСИИ, 2) & vbCrLf & vbCrLf & _
              "Обновиться? Рабочие данные не пострадают — заменяется только программа.", _
              vbYesNo + vbQuestion, "Обновление") = vbYes Then
        УстановитьОбновление папка, номерВСети
    End If
End Sub


Private Sub УстановитьОбновление(папка As String, номер As String)
    Dim источник As String
    Dim врем As String
    Dim ожидаемая As String
    Dim фактическая As String
    Dim архив As String

    источник = папка & ФАЙЛ_ПРОГРАММЫ
    If Not ФайлСуществует(источник) Then
        MsgBox "Файл версии есть, а самой программы в общей папке нет:" & vbCrLf & источник, _
               vbExclamation, "Обновление"
        Exit Sub
    End If

    On Error GoTo Сбой
    ' Сначала копируем во временный файл и проверяем — работающую
    ' программу заменяем только тем, что заведомо доехало целиком.
    врем = ThisWorkbook.Path & Application.PathSeparator & "~обновление.xlsm"
    If ФайлСуществует(врем) Then Kill врем
    FileCopy источник, врем

    ожидаемая = ПолеВерсии(папка & ФАЙЛ_ВЕРСИИ, 4)
    If ожидаемая <> "" Then
        фактическая = SHA256(врем)
        If StrComp(фактическая, ожидаемая, vbTextCompare) <> 0 Then
            Kill врем
            MsgBox "Обновление отменено: контрольная сумма не совпала." & vbCrLf & vbCrLf & _
                   "Ожидалась: " & ожидаемая & vbCrLf & _
                   "Получилась: " & фактическая & vbCrLf & vbCrLf & _
                   "Файл скопировался не полностью или его подменили. " & _
                   "Старая версия осталась на месте — работать можно. " & _
                   "Скажите мастеру, чтобы опубликовал версию заново.", _
                   vbCritical, "Обновление"
            Exit Sub
        End If
    End If

    ' Предыдущая версия не удаляется: откат — это переименование файла,
    ' а не восстановление из резервной копии.
    архив = Путь("Папка программы", "") & "Версии" & Application.PathSeparator
    MkPath архив
    FileCopy ThisWorkbook.FullName, _
             архив & "ЛНК_Программа_v" & Настройка("Версия программы") & ".xlsm"

    ЗаписатьНастройку "Версия программы", номер
    Данные().Save

    MsgBox "Файл проверен и готов к установке." & vbCrLf & vbCrLf & _
           "Excel не может заменить книгу, которая сейчас открыта. Поэтому:" & vbCrLf & _
           "1. Закройте Excel полностью." & vbCrLf & _
           "2. В папке " & ThisWorkbook.Path & " удалите " & ФАЙЛ_ПРОГРАММЫ & vbCrLf & _
           "3. Переименуйте ~обновление.xlsm в " & ФАЙЛ_ПРОГРАММЫ & vbCrLf & _
           "4. Откройте её." & vbCrLf & vbCrLf & _
           "Предыдущая версия сохранена в папке «Версии» — если новая чем-то " & _
           "не устроит, откат это просто переименование файла.", _
           vbInformation, "Обновление"
    Exit Sub

Сбой:
    On Error Resume Next
    If ФайлСуществует(врем) Then Kill врем
    MsgBox "Обновление не установлено." & vbCrLf & "Ошибка: " & Err.Description & vbCrLf & vbCrLf & _
           "Старая версия осталась на месте, работать можно.", vbCritical, "Обновление"
End Sub


' ── Публикация версии (мастер) ────────────────────────────────────────

Public Sub ОпубликоватьВерсию()
    Dim папка As String
    Dim номер As String
    Dim f As Integer
    Dim сумма As String

    If Not ОбщаяПапкаЕсть() Then Exit Sub
    номер = Trim$(InputBox("Номер публикуемой версии (например 1.1):", _
                           "Публикация версии", СледующийНомер()))
    If номер = "" Then Exit Sub

    On Error GoTo Сбой
    папка = ПапкаОбмена() & "Версия" & Application.PathSeparator
    MkPath папка

    ThisWorkbook.Save
    If ФайлСуществует(папка & ФАЙЛ_ПРОГРАММЫ) Then Kill папка & ФАЙЛ_ПРОГРАММЫ
    FileCopy ThisWorkbook.FullName, папка & ФАЙЛ_ПРОГРАММЫ
    сумма = SHA256(папка & ФАЙЛ_ПРОГРАММЫ)

    f = FreeFile
    Open папка & ФАЙЛ_ВЕРСИИ For Output As #f
    Print #f, номер
    Print #f, Format$(Now, "dd.mm.yyyy hh:nn")
    Print #f, ТекущийСотрудник() & " (" & КодРабочегоМеста() & ")"
    Print #f, сумма
    Close #f

    ЗаписатьНастройку "Версия программы", номер

    MsgBox "Версия " & номер & " опубликована." & vbCrLf & _
           "Контрольная сумма: " & Left$(сумма, 16) & "…" & vbCrLf & vbCrLf & _
           "Остальные рабочие места предложат обновиться при следующем запуске.", _
           vbInformation, "Публикация версии"
    Exit Sub
Сбой:
    On Error Resume Next
    Close #f
    MsgBox "Не удалось опубликовать версию." & vbCrLf & Err.Description, _
           vbCritical, "Публикация версии"
End Sub


' ── Контрольная сумма ─────────────────────────────────────────────────

' certutil — штатная утилита Windows, прав администратора не требует.
' Её вывод: строка заголовка, строка с суммой, строка «успешно».
Public Function SHA256(путь As String) As String
    Dim wsh As Object
    Dim врем As String
    Dim f As Integer
    Dim строка As String
    Dim результат As String

    On Error GoTo Сбой
    врем = Environ$("TEMP") & "\лнк_сумма.txt"
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "cmd /c certutil -hashfile """ & путь & """ SHA256 > """ & врем & """", 0, True

    If Not ФайлСуществует(врем) Then Exit Function
    f = FreeFile
    Open врем For Input As #f
    Do While Not EOF(f)
        Line Input #f, строка
        строка = Replace(Trim$(строка), " ", "")
        ' Сумма — единственная строка из 64 шестнадцатеричных знаков.
        If Len(строка) = 64 Then
            If ЭтоHex(строка) Then результат = LCase$(строка)
        End If
    Loop
    Close #f
    Kill врем
    SHA256 = результат
    Exit Function
Сбой:
    On Error Resume Next
    Close #f
End Function


Private Function ЭтоHex(s As String) As Boolean
    Dim i As Long
    Dim c As String
    For i = 1 To Len(s)
        c = LCase$(Mid$(s, i, 1))
        If InStr("0123456789abcdef", c) = 0 Then Exit Function
    Next i
    ЭтоHex = True
End Function


' ── Вспомогательное ───────────────────────────────────────────────────

Private Function ПолеВерсии(файл As String, номерСтроки As Long) As String
    Dim f As Integer
    Dim i As Long
    Dim s As String

    If Not ФайлСуществует(файл) Then Exit Function
    On Error Resume Next
    f = FreeFile
    Open файл For Input As #f
    For i = 1 To номерСтроки
        If EOF(f) Then Exit For
        Line Input #f, s
    Next i
    Close #f
    ПолеВерсии = Trim$(s)
End Function


' Версии вида 1.10 и 1.9: посегментное сравнение, а не строковое —
' строковое поставило бы 1.10 раньше 1.9.
Private Function СравнитьВерсии(a As String, b As String) As Long
    Dim ча() As String
    Dim чб() As String
    Dim i As Long
    Dim x As Long, y As Long

    ча = Split(a & ".0.0", ".")
    чб = Split(b & ".0.0", ".")
    For i = 0 To 2
        x = CLng(Val(ча(i)))
        y = CLng(Val(чб(i)))
        If x > y Then
            СравнитьВерсии = 1
            Exit Function
        ElseIf x < y Then
            СравнитьВерсии = -1
            Exit Function
        End If
    Next i
End Function


Private Function СледующийНомер() As String
    Dim текущая As String
    Dim части() As String
    текущая = Настройка("Версия программы")
    If текущая = "" Then
        СледующийНомер = "1.0"
        Exit Function
    End If
    части = Split(текущая & ".0", ".")
    СледующийНомер = части(0) & "." & CStr(CLng(Val(части(1))) + 1)
End Function


Private Function ОбщаяПапкаЕсть() As Boolean
    If ПапкаСуществует(ПапкаОбмена()) Then
        ОбщаяПапкаЕсть = True
        Exit Function
    End If
    MsgBox "Общая папка недоступна:" & vbCrLf & ПапкаОбмена(), vbExclamation, "Обновление"
End Function
