Attribute VB_Name = "modConfig"
Option Explicit

' Номера столбцов листа "Журнал" (см. заголовки строки 4 на самом листе)
Public Const COL_METHOD As Long = 2      ' Метод контроля
Public Const COL_YEAR As Long = 3        ' Год
Public Const COL_SERIES As Long = 4      ' № в серии (формула)
Public Const COL_REGNUM As Long = 5      ' Регистрационный номер (формула)
Public Const COL_DATE As Long = 6        ' Дата регистрации
Public Const COL_CUSTOMER As Long = 7    ' Заказывающее подразделение
Public Const COL_OBJECT As Long = 8      ' Наименование объекта контроля
Public Const COL_EXECUTOR As Long = 9    ' Исполнитель (ФИО)
Public Const COL_GENERATED As Long = 14  ' Документ сформирован (дата, макросом)

' Пути — всегда относительно расположения этой книги, чтобы работало
' и на общей сетевой папке, и на любом ПК без правки путей в коде.
Public Function TemplatesFolder() As String
    TemplatesFolder = ThisWorkbook.Path & "\Шаблоны\"
End Function

Public Function OutputFolder() As String
    Dim p As String
    p = ThisWorkbook.Path & "\Готовые\"
    If Dir(p, vbDirectory) = "" Then MkDir p
    OutputFolder = p
End Function

Public Function TemplateFileFor(methodName As String) As String
    Select Case methodName
        Case "ВИК": TemplateFileFor = "Акт_ВИК.docx"
        Case "УЗТ": TemplateFileFor = "Протокол_УЗТ.docx"
        Case "УЗД": TemplateFileFor = "Заключение_УЗД.docx"
        Case "ТВ":  TemplateFileFor = "Протокол_ТВ.docx"
        Case "АЭ":  TemplateFileFor = "Заключение_АЭ.docx"
        Case "ТК":  TemplateFileFor = "Протокол_ТК.docx"
        Case "ВД":  TemplateFileFor = "Протокол_ВД.docx"
        Case Else:  TemplateFileFor = ""   ' МК — шаблона пока нет, добавите по образцу остальных
    End Select
End Function

Public Sub MkPath(p As String)
    Dim parts() As String, cur As String, i As Long
    parts = Split(p, "\")
    cur = parts(0)
    For i = 1 To UBound(parts)
        If parts(i) = "" Then Exit For
        cur = cur & "\" & parts(i)
        If Dir(cur, vbDirectory) = "" Then MkDir cur
    Next i
End Sub

Public Function ОчиститьДляИмениФайла(s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|", Chr(10), Chr(13))
    ОчиститьДляИмениФайла = s
    For i = LBound(bad) To UBound(bad)
        ОчиститьДляИмениФайла = Replace(ОчиститьДляИмениФайла, bad(i), " ")
    Next i
    Do While InStr(ОчиститьДляИмениФайла, "  ") > 0
        ОчиститьДляИмениФайла = Replace(ОчиститьДляИмениФайла, "  ", " ")
    Loop
    ОчиститьДляИмениФайла = Trim(ОчиститьДляИмениФайла)
    If Len(ОчиститьДляИмениФайла) > 60 Then ОчиститьДляИмениФайла = Left(ОчиститьДляИмениФайла, 60)
End Function

Public Function ПутьКДокументу(ws As Worksheet, r As Long) As String
    Dim methodName As String, objName As String, folder As String
    methodName = Trim(ws.Cells(r, COL_METHOD).Value)
    objName = Trim(ws.Cells(r, COL_OBJECT).Value)
    folder = OutputFolder() & ws.Cells(r, COL_YEAR).Value & "\" & methodName & "\"
    ПутьКДокументу = folder & Replace(ws.Cells(r, COL_REGNUM).Text, "/", "-") & " " & _
                     ОчиститьДляИмениФайла(objName) & ".docx"
End Function
