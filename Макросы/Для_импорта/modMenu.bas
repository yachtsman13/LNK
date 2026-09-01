Attribute VB_Name = "modMenu"
Option Explicit

' =======================================================================
'  Меню: единственный лист, с которого начинается работа.
'
'  Пункты — гиперссылки, а не кнопки-картинки: гиперссылка переживает
'  копирование книги между машинами, а кнопка на форме — не всегда.
' =======================================================================

Private Const Л_МЕНЮ As String = "Меню"


Public Sub ПостроитьМеню()
    Dim ws As Worksheet
    Dim r As Long

    Application.ScreenUpdating = False
    Set ws = ПересоздатьЛистМеню()

    ws.Range("B2").Value = "ЛАБОРАТОРИЯ НЕРАЗРУШАЮЩЕГО КОНТРОЛЯ"
    ws.Range("B3").Value = "Рабочее место " & КодРабочегоМеста() & _
                           IIf(ТекущийСотрудник() = "", "", " · " & ТекущийСотрудник()) & _
                           IIf(ЭтоМастер(), " · мастер", "")
    ws.Range("B2").Font.Size = 18
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(31, 56, 100)
    ws.Range("B3").Font.Italic = True
    ws.Range("B3").Font.Color = RGB(89, 89, 89)

    r = 5
    r = Раздел(ws, r, "Сегодня")
    r = Пункт(ws, r, "Обновить Дашборд", _
              "Что просрочено, что истекает, что ждёт сборки")
    r = ПунктЛист(ws, r, "Дашборд", "Дашборд", "Открыть уже построенный")
    r = r + 1

    r = Раздел(ws, r, "Ежедневная работа")
    r = Пункт(ws, r, "Новая запись в журнале", _
              "Заводит черновую строку: номера будут выданы при регистрации")
    r = Пункт(ws, r, "Зарегистрировать запись", _
              "Выдаёт настоящие номера сразу по всем методам строки")
    r = Пункт(ws, r, "Сформировать документы", _
              "Акты и протоколы Word по всем методам текущей строки")
    r = Пункт(ws, r, "Аннулировать заключение", _
              "По итогам рекламации: строка остаётся, номер не переиспользуется")
    r = r + 1

    r = Раздел(ws, r, "Данные")
    r = ПунктЛист(ws, r, "Журнал", Л_ЖУРНАЛ, "Один объект за дату, методы колонками")
    r = ПунктЛист(ws, r, "Объекты контроля", Л_ОБЪЕКТЫ, "Справочник, история работ по объекту")
    r = ПунктЛист(ws, r, "Средства измерений", Л_СИ, "Поверка, характеристики, обслуживание")
    r = ПунктЛист(ws, r, "Специалисты", Л_СПЕЦИАЛИСТЫ, "Кадры")
    r = ПунктЛист(ws, r, "Аттестации", Л_АТТЕСТАЦИИ, "Удостоверения и сроки по СДАНК-02-2020")
    r = ПунктЛист(ws, r, "Кто где", Л_КТО_ГДЕ, "Отпуска, больничные, выезды")
    r = r + 1

    r = Раздел(ws, r, "Обмен с другими рабочими местами")
    r = Пункт(ws, r, "Отправить мои изменения", _
              "Всё, что изменилось с прошлой отправки, уходит в общую папку")
    r = Пункт(ws, r, "Собрать свод", _
              "Действие мастера: применяет все дельты и публикует общую версию")
    r = Пункт(ws, r, "Получить свод", _
              "Заменить локальные данные версией мастера")
    r = Пункт(ws, r, "Взять роль мастера на себя", _
              "Мастером может быть любой; двух одновременно не бывает")
    r = Пункт(ws, r, "Зарегистрировать носитель", _
              "Привязка штатной флешки по серийному номеру тома")
    r = r + 1

    r = Раздел(ws, r, "Печать и закрытие года")
    r = Пункт(ws, r, "Построить печатные журналы", _
              "Перестраивает листы Печать_* по методам из журнала")
    r = r + 1

    r = Раздел(ws, r, "Служебное")
    r = Пункт(ws, r, "Сделать резервную копию", _
              "Копия данных в папку «Копии»; то же делается при закрытии книги")
    r = Пункт(ws, r, "Пересчитать счётчики по журналу", _
              "Стартовые значения при переходе и восстановление после сбоя")
    r = ПунктЛист(ws, r, "Настройки", Л_НАСТРОЙКИ, "Код рабочего места, пути, мастер")
    r = ПунктЛист(ws, r, "Сверка миграции", "Сверка_миграции", _
                  "Что при переносе данных требует вашего решения")

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 42
    ws.Columns("C").ColumnWidth = 72
    ws.Activate
    ws.Range("A1").Select
    Application.ScreenUpdating = True
End Sub


Private Function ПересоздатьЛистМеню() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(Л_МЕНЮ).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = Л_МЕНЮ
    ws.Tab.Color = RGB(31, 56, 100)
    ws.Cells.Interior.Color = RGB(255, 255, 255)
    ws.Cells.Font.Name = "Calibri"
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    Set ПересоздатьЛистМеню = ws
End Function


Private Function Раздел(ws As Worksheet, r As Long, текст As String) As Long
    ws.Cells(r, 2).Value = текст
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 3))
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    ws.Rows(r).RowHeight = 20
    Раздел = r + 1
End Function


Private Function Пункт(ws As Worksheet, r As Long, текст As String, _
                       подсказка As String) As Long
    ws.Cells(r, 2).Value = "»  " & текст
    ws.Cells(r, 2).Font.Color = RGB(0, 51, 160)
    ws.Cells(r, 2).Font.Underline = xlUnderlineStyleSingle
    ws.Cells(r, 3).Value = подсказка
    ws.Cells(r, 3).Font.Color = RGB(89, 89, 89)
    ws.Cells(r, 3).Font.Italic = True
    ws.Rows(r).RowHeight = 18
    Пункт = r + 1
End Function


Private Function ПунктЛист(ws As Worksheet, r As Long, текст As String, _
                           лист As String, подсказка As String) As Long
    On Error Resume Next
    ws.Hyperlinks.Add Anchor:=ws.Cells(r, 2), Address:="", _
                      SubAddress:="'" & лист & "'!A1", TextToDisplay:="»  " & текст
    ws.Cells(r, 3).Value = подсказка
    ws.Cells(r, 3).Font.Color = RGB(89, 89, 89)
    ws.Cells(r, 3).Font.Italic = True
    ws.Rows(r).RowHeight = 18
    ПунктЛист = r + 1
End Function


' Пункты-макросы вызываются двойным щелчком: обработчик в модуле листа
' «Меню» разбирает, на какой строке щёлкнули. Гиперссылка на макрос в
' Excel не создаётся, а кнопки-картинки переживают перенос книги хуже.
Public Sub ЗапуститьПунктМеню(текст As String)
    Dim имя As String
    имя = Replace(Replace(текст, "»", ""), Chr(160), " ")
    имя = Trim$(имя)

    Select Case имя
        Case "Обновить Дашборд": modDashboard.ПостроитьДашборд
        Case "Новая запись в журнале": modJournal.НоваяЗапись
        Case "Зарегистрировать запись": modJournal.ЗарегистрироватьЗапись
        Case "Сформировать документы": modGenerate.СформироватьДокументы
        Case "Аннулировать заключение": modJournal.АннулироватьЗапись
        Case "Построить печатные журналы": modPrintJournals.ПостроитьВсеПечатныеЖурналы
        Case "Сделать резервную копию": modBackup.СделатьЛокальнуюКопию
        Case "Пересчитать счётчики по журналу": modBackup.ВосстановитьСчётчикиИзЖурнала
        Case "Отправить мои изменения": modExchange.ОтправитьМоиИзменения
        Case "Собрать свод": modExchange.СобратьСвод
        Case "Получить свод": modExchange.ПолучитьСвод
        Case "Взять роль мастера на себя": modExchange.ВзятьРольМастера
        Case "Зарегистрировать носитель": modMedia.ЗарегистрироватьНоситель
    End Select
End Sub
