Attribute VB_Name = "modAccessChecks"
Option Explicit
Option Compare Database

' Проверяет по таблице tblАттестации, действительна ли аттестация исполнителя
' по нужному методу на дату контроля. Возвращает "" если всё в порядке, иначе —
' текст предупреждения для показа перед формированием документа.
Public Function ПроверитьАттестацию(fio As String, methodName As String, onDate As Variant) As String
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim checkDate As Date
    Dim found As Boolean

    If IsDate(onDate) Then checkDate = CDate(onDate) Else checkDate = Date

    sql = "SELECT СрокДействия FROM tblАттестации WHERE ФИО = " & SqlQuote(fio) & _
          " AND Метод = " & SqlQuote(methodName)
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)

    found = Not rs.EOF
    Do While Not rs.EOF
        If Not IsNull(rs!СрокДействия) Then
            If CDate(rs!СрокДействия) < checkDate Then
                ПроверитьАттестацию = "Удостоверение «" & fio & "» по методу " & methodName & _
                    " просрочено (действительно до " & Format(rs!СрокДействия, "dd.mm.yyyy") & ")."
                rs.Close
                Exit Function
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close

    If Not found Then
        ПроверитьАттестацию = "В таблице «tblАттестации» не нашлась запись «" & fio & "» по методу " & _
            methodName & " — проверьте квалификацию вручную перед выдачей документа."
    End If
End Function

Private Function SqlQuote(s As String) As String
    SqlQuote = "'" & Replace(s, "'", "''") & "'"
End Function
