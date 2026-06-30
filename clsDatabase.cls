VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Conn As Object
Private m_QueryDefs As Collection
Private m_Relations As Collection
Private m_TableDefs As Collection

Public RecordsAffected As Long

Private Sub Class_Initialize()
    Set Conn = CreateObject("ADODB.Connection")
    Set m_QueryDefs = New Collection
    Set m_Relations = New Collection
    Set m_TableDefs = New Collection
End Sub

Public Function Connect(server As String, db As String, uID As String, PWD As String) As Boolean
    On Error Resume Next
    Conn.Open "Provider=SQLOLEDB;Data Source=" & server & ";Initial Catalog=" & db & ";User ID=" & uID & ";Password=" & PWD & ";"
    Connect = (Err.number = 0)
End Function

Public Sub ExecuteSQL(SQL As String, Optional Options As Variant = dbFailOnError)
    If Left(SQL, 5) <> "ALTER" Then
        Debug.Print "ExecuteSQL     " & SQL
    End If
    On Error GoTo Loi
    Conn.Execute SQL, RecordsAffected
    Exit Sub
Loi:     If Left(SQL, 5) <> "ALTER" Then

        Debug.Print Err.number, "clsDatabase.ExecuteSQL", Err.Description
    End If
End Sub

Private Function ConvertSQL(ByVal SQL As String) As String
    Dim result As String
    result = SQL
    result = FixConvertError(result)
    result = FixPlusConcat(result)
    result = FixElseEndConcat(result)
    result = Replace(result, "DISTINCTROW", "DISTINCT")
    result = Replace(result, "distinctrow", "DISTINCT")
    result = ConvertDateDelimiters(result)
    result = ConvertWildcards(result)
    result = ConvertIIFToCase(result)
    result = Replace(result, "Format(", "CONVERT(varchar, ")
    result = FixFormatDate(result)
    result = Replace(result, "Fix(", "ROUND(")
    result = Replace(result, "Date()", "GETDATE()")
    result = Replace(result, "Now()", "GETDATE()")
    ConvertSQL = result
End Function

Private Function FixFormatDate(ByVal SQL As String) As String
    Dim result As String
    result = SQL
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.pattern = "Format\(([^,]+),\s*['""]([^'""]+)['""]\)"
    Dim matches As Object
    Set matches = regex.Execute(result)
    Dim match As Object
    Dim fieldName As String
    Dim dateFormat As String
    Dim sqlCode As String
    For Each match In matches
        fieldName = match.SubMatches(0)
        dateFormat = LCase(match.SubMatches(1))
        Select Case dateFormat
            Case "dd/mm/yy", "dd/mm/yyyy", "dd\/mm\/yy", "dd\/mm\/yyyy"
                sqlCode = "CONVERT(varchar, " & fieldName & ", 3)"
            Case "mm/dd/yy", "mm/dd/yyyy"
                sqlCode = "CONVERT(varchar, " & fieldName & ", 1)"
            Case "yyyymmdd"
                sqlCode = "CONVERT(varchar, " & fieldName & ", 112)"
            Case "dd/mm/yyyy hh:nn:ss", "dd/mm/yy hh:nn:ss"
                sqlCode = "CONVERT(varchar, " & fieldName & ", 103)"
            Case Else
                sqlCode = "CONVERT(varchar, " & fieldName & ", 103)"
        End Select
        result = Replace(result, match.Value, sqlCode)
    Next
    FixFormatDate = result
End Function

Private Function FixConvertError(ByVal SQL As String) As String
    Dim result As String
    result = SQL
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.pattern = "CONVERT\s*\(\s*varchar\s*,\s*([^,]+)\s*,\s*['""]([^'""]+)['""]\s*\)"
    Dim matches As Object
    Set matches = regex.Execute(result)
    Dim match As Object
    Dim fieldName As String
    Dim dateFormat As String
    Dim formatCode As String
    For Each match In matches
        fieldName = match.SubMatches(0)
        dateFormat = LCase(match.SubMatches(1))
        Select Case dateFormat
        Case "dd/mm/yy", "dd/mm/yyyy", "dd\/mm\/yy", "dd\/mm\/yyyy"
            formatCode = "3"
        Case "mm/dd/yy", "mm/dd/yyyy"
            formatCode = "1"
        Case "yyyymmdd"
            formatCode = "112"
        Case Else
            formatCode = "3"
        End Select
        result = Replace(result, match.Value, "CONVERT(varchar, " & fieldName & ", " & formatCode & ")")
    Next
    FixConvertError = result
End Function

Private Function FixElseEndConcat(ByVal SQL As String) As String
    Dim result As String
    result = SQL
    Dim pos As Integer
    pos = InStr(1, result, "ELSE '' END+", vbTextCompare)
    Do While pos > 0
        Dim startField As Integer
        startField = pos + 11
        Dim endField As Integer
        endField = startField
        Do While endField <= Len(result) And (Mid(result, endField, 1) Like "[A-Za-z0-9_.]")
            endField = endField + 1
        Loop
        Dim fieldName As String
        fieldName = Mid(result, startField, endField - startField)
        Dim newText As String
        newText = "ELSE ISNULL(" & fieldName & ", '') END"
        result = Left(result, pos - 1) & newText & Mid(result, endField)
        pos = InStr(pos + 1, result, "ELSE '' END+", vbTextCompare)
    Loop
    FixElseEndConcat = result
End Function

Private Function FixPlusConcat(ByVal SQL As String) As String
    Dim result As String
    result = SQL
    Dim pos As Integer
    pos = InStr(1, result, "+ '", vbTextCompare)
    If pos > 0 Then
        Dim leftStart As Integer
        leftStart = pos - 1
        Do While leftStart > 1 And Mid(result, leftStart, 1) = " "
            leftStart = leftStart - 1
        Loop
        Do While leftStart > 1 And (Mid(result, leftStart, 1) Like "[A-Za-z0-9_.]")
            leftStart = leftStart - 1
        Loop
        Dim leftField As String
        leftField = Trim(Mid(result, leftStart + 1, pos - leftStart - 1))
        Dim quote1 As Integer
        quote1 = InStr(pos, result, "'")
        Dim quote2 As Integer
        quote2 = InStr(quote1 + 1, result, "'")
        Dim separator As String
        separator = Mid(result, quote1, quote2 - quote1 + 1)
        Dim rightPos As Integer
        rightPos = InStr(quote2, result, "+")
        If rightPos > 0 Then
            Dim rightStart As Integer
            rightStart = rightPos + 1
            Do While rightStart <= Len(result) And Mid(result, rightStart, 1) = " "
                rightStart = rightStart + 1
            Loop
            Dim rightEnd As Integer
            rightEnd = rightStart
            Do While rightEnd <= Len(result) And (Mid(result, rightEnd, 1) Like "[A-Za-z0-9_.]")
                rightEnd = rightEnd + 1
            Loop
            Dim rightField As String
            rightField = Trim(Mid(result, rightStart, rightEnd - rightStart))
            Dim newConcat As String
            newConcat = "CONCAT(" & leftField & ", " & separator & ", ISNULL(" & rightField & ", ''))"
            result = Replace(result, leftField & " + " & separator & " + " & rightField, newConcat)
        End If
    End If
    FixPlusConcat = result
End Function

Private Function ConvertDateDelimiters(ByVal SQL As String) As String
    Dim result As String
    Dim i As Integer
    Dim inString As Boolean
    Dim startPos As Integer
    Dim dateStr As String
    Dim ch As String
    result = SQL
    i = 1
    inString = False
    Do While i <= Len(result)
        ch = Mid(result, i, 1)
        If ch = "'" Then
            inString = Not inString
        End If
        If Not inString And ch = "#" Then
            startPos = i
            i = i + 1
            dateStr = ""
            Do While i <= Len(result)
                ch = Mid(result, i, 1)
                If ch = "#" Then Exit Do
                dateStr = dateStr & ch
                i = i + 1
            Loop
            dateStr = ConvertDateFormat(dateStr)
            result = Left(result, startPos - 1) & "'" & dateStr & "'" & Mid(result, i + 1)
            i = startPos + Len(dateStr) + 2
        End If
        i = i + 1
    Loop
    ConvertDateDelimiters = result
End Function

Private Function ConvertDateFormat(ByVal dateStr As String) As String
    Dim parts() As String
    Dim day As String
    Dim month As String
    Dim year As String
    dateStr = Trim(dateStr)
    If InStr(dateStr, "/") > 0 Then
        parts = Split(dateStr, "/")
        If UBound(parts) = 2 Then
            month = parts(0)
            day = parts(1)
            year = parts(2)
            If Len(year) = 2 Then
                If CInt(year) >= 30 Then
                    year = "19" & year
                Else
                    year = "20" & year
                End If
            End If
            ConvertDateFormat = year & "-" & Right("0" & month, 2) & "-" & Right("0" & day, 2)
            Exit Function
        End If
    End If
    If InStr(dateStr, "-") > 0 Then
        parts = Split(dateStr, "-")
        If UBound(parts) = 2 Then
            day = parts(0)
            month = parts(1)
            year = parts(2)
            If Len(year) = 2 Then
                If CInt(year) >= 30 Then
                    year = "19" & year
                Else
                    year = "20" & year
                End If
            End If
            ConvertDateFormat = year & "-" & Right("0" & month, 2) & "-" & Right("0" & day, 2)
            Exit Function
        End If
    End If
    ConvertDateFormat = dateStr
End Function

Private Function ConvertIIFToCase(ByVal SQL As String) As String
    Dim result As String
    Dim pos As Integer
    Dim i As Integer
    Dim level As Integer
    Dim startPos As Integer
    Dim endPos As Integer
    Dim commaPos1 As Integer
    Dim commaPos2 As Integer
    Dim condition As String
    Dim truePart As String
    Dim falsePart As String
    result = SQL
    pos = InStr(1, result, "IIF(", vbTextCompare)
    Do While pos > 0
        startPos = pos
        level = 1
        i = pos + 4
        Do While i <= Len(result)
            Dim ch As String
            ch = Mid(result, i, 1)
            If ch = "(" Then level = level + 1
            If ch = ")" Then
                level = level - 1
                If level = 0 Then
                    endPos = i
                    Exit Do
                End If
            End If
            i = i + 1
        Loop
        commaPos1 = 0
        level = 0
        For i = pos + 4 To endPos - 1
            ch = Mid(result, i, 1)
            If ch = "(" Then level = level + 1
            If ch = ")" Then level = level - 1
            If ch = "," And level = 0 Then
                commaPos1 = i
                Exit For
            End If
        Next
        commaPos2 = 0
        If commaPos1 > 0 Then
            level = 0
            For i = commaPos1 + 1 To endPos - 1
                ch = Mid(result, i, 1)
                If ch = "(" Then level = level + 1
                If ch = ")" Then level = level - 1
                If ch = "," And level = 0 Then
                    commaPos2 = i
                    Exit For
                End If
            Next
        End If
        If commaPos1 > 0 And commaPos2 > 0 Then
            condition = Trim(Mid(result, pos + 4, commaPos1 - pos - 4))
            truePart = Trim(Mid(result, commaPos1 + 1, commaPos2 - commaPos1 - 1))
            falsePart = Trim(Mid(result, commaPos2 + 1, endPos - commaPos2 - 1))
            Dim caseExpr As String
            caseExpr = "CASE WHEN " & condition & " THEN " & truePart & " ELSE " & falsePart & " END"
            result = Left(result, startPos - 1) & caseExpr & Mid(result, endPos + 1)
        End If
        pos = InStr(1, result, "IIF(", vbTextCompare)
    Loop
    ConvertIIFToCase = result
End Function

Private Function ConvertWildcards(ByVal SQL As String) As String
    Dim result As String
    Dim i As Long
    Dim inString As Boolean
    Dim inLike As Boolean
    Dim patternStart As Long
    Dim pattern As String
    Dim newPattern As String
    Dim ch As String
    Dim j As Long
    Dim inBracket As Boolean
    Dim bracketContent As String
    result = SQL
    i = 1
    inString = False
    inLike = False
    Do While i <= Len(result)
        ch = Mid(result, i, 1)
        If ch = "'" Then
            inString = Not inString
        End If
        If Not inString Then
            If i + 4 <= Len(result) Then
                If UCase(Mid(result, i, 5)) = " LIKE" Then
                    inLike = True
                    i = i + 5
                    GoTo NextIteration
                End If
            End If
        End If
        If inLike And Not inString And ch = "'" Then
            patternStart = i + 1
            pattern = ""
            newPattern = ""
            i = i + 1
            Do While i <= Len(result)
                ch = Mid(result, i, 1)
                If ch = "'" Then Exit Do
                pattern = pattern & ch
                i = i + 1
            Loop
            j = 1
            inBracket = False
            Do While j <= Len(pattern)
                ch = Mid(pattern, j, 1)
                If ch = "[" Then
                    inBracket = True
                    bracketContent = "["
                ElseIf inBracket And ch = "]" Then
                    inBracket = False
                    newPattern = newPattern & bracketContent & "]"
                ElseIf inBracket Then
                    bracketContent = bracketContent & ch
                Else
                    Select Case ch
                    Case "*"
                        newPattern = newPattern & "%"
                    Case "?"
                        newPattern = newPattern & "_"
                    Case "#"
                        newPattern = newPattern & "[0-9]"
                    Case "!"
                        If j > 1 And Mid(pattern, j - 1, 1) = "[" Then
                            newPattern = newPattern & "^"
                        Else
                            newPattern = newPattern & ch
                        End If
                    Case Else
                        newPattern = newPattern & ch
                    End Select
                End If
                j = j + 1
            Loop
            result = Left(result, patternStart - 1) & newPattern & Mid(result, i)
            inLike = False
        End If
NextIteration:
        i = i + 1
    Loop
    ConvertWildcards = result
End Function
Public Function FixSQL(ByVal SQL As String) As String

    Dim s As String
    s = SQL

    ' DISTINCTROW -> DISTINCT
    s = Replace(s, "DISTINCTROW", "DISTINCT", , , vbTextCompare)

    ' SELECTROW -> SELECT
    s = Replace(s, "SELECTROW", "SELECT", , , vbTextCompare)

    ' WHERE TRUE
    s = Replace(s, "(TRUE)", "(1=1)", , , vbTextCompare)
    s = Replace(s, " TRUE ", " 1=1 ", , , vbTextCompare)

    ' Val(field) -> CAST(field AS BIGINT) - S?A L?I PATTERN
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    reg.Global = True
    reg.IgnoreCase = True
    ' Pattern m?i: b?t du?c val(ChungTu.SoHieu), val(ChungTuP.SoHieu), val(SoHieu), v.v.
    reg.pattern = "Val\s*\(\s*([\w\.]+\.)?SoHieu\s*\)"
    
    s = reg.Replace(s, "TRY_CAST($1SoHieu AS BIGINT)")
    
    ' N?u v?n mu?n dùng CAST thông thu?ng:
    ' s = reg.Replace(s, "CAST($1SoHieu AS BIGINT)")

    ' Lo?i b? DISTINCT n?u có GROUP BY
    If InStr(1, s, "GROUP BY", vbTextCompare) > 0 Then
        s = Replace(s, "SELECT DISTINCT", "SELECT", , , vbTextCompare)
    End If

    FixSQL = s

End Function
Public Function OpenRecordset(Source As String, Optional Typ As Variant = 2, Optional Typ2 As Variant = 2) As Object

    Dim rs As Object
    Dim SQL As String
    SQL = Source
    SQL = FixSQL(SQL)
    SQL = FixSpecialSQL(SQL)
    SQL = ConvertSQL(SQL)
    SQL = Replace(SQL, "DISTINCTROW", "DISTINCT")
    SQL = Replace(SQL, "distinctrow", "DISTINCT")
    SQL = Trim(SQL)
    On Error Resume Next
    SQL = m_QueryDefs(Source)
    On Error GoTo 0
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3
    rs.CursorType = 3
    rs.LockType = 1
    Dim rsType As Integer
    Select Case Typ
    Case 2:
        rsType = 3
    Case Else:
        rsType = 3
    End Select
    Debug.Print "OpenRecordset  " & SQL
    rs.Open SQL, Conn, rsType, 3
    Set OpenRecordset = rs
End Function

Public Function FixSpecialSQL(ByVal SQL As String) As String
    Dim result As String
    result = SQL
    
    ' 1. S?a CONVERT(varchar, NgayKy,'dd/mm/yy') -> CONVERT(varchar, NgayKy, 103)
    Dim pos As Integer
    pos = InStr(1, result, "CONVERT(varchar,", vbTextCompare)
    Do While pos > 0
        Dim comma1 As Integer, comma2 As Integer, endPos As Integer
        comma1 = InStr(pos + 16, result, ",")
        If comma1 > 0 Then
            comma2 = InStr(comma1 + 1, result, ",")
            If comma2 > 0 Then
                endPos = InStr(comma2 + 1, result, ")")
                If endPos > 0 Then
                    Dim fieldName As String
                    fieldName = Trim(Mid(result, pos + 16, comma1 - pos - 16))
                    Dim newConvert As String
                    newConvert = "CONVERT(varchar, " & fieldName & ", 103)"
                    result = Left(result, pos - 1) & newConvert & Mid(result, endPos + 1)
                End If
            End If
        End If
        pos = InStr(pos + 1, result, "CONVERT(varchar,", vbTextCompare)
    Loop  ' <-- THI?U CÁI NÀY
    
    ' 2. S?a ELSE '' END+DienGiai -> ELSE ISNULL(DienGiai, '') END
    pos = InStr(1, result, "ELSE '' END+", vbTextCompare)
    Do While pos > 0
        Dim startField As Integer
        startField = pos + 11
        Dim endField As Integer
        endField = startField
        Do While endField <= Len(result) And (Mid(result, endField, 1) Like "[A-Za-z0-9_.]")
            endField = endField + 1
        Loop
        Dim fieldName2 As String
        fieldName2 = Mid(result, startField, endField - startField)
        Dim newText As String
        newText = "ELSE ISNULL(" & fieldName2 & ", '') END"
        result = Left(result, pos - 1) & newText & Mid(result, endField)
        pos = InStr(pos + 1, result, "ELSE '' END+", vbTextCompare)
    Loop  ' <-- Cái này b?n có r?i
    
    ' 3. S?a THEN field+' - ' + field -> CONCAT
    pos = InStr(1, result, "THEN ", vbTextCompare)
    Do While pos > 0
        Dim plusPos As Integer
        plusPos = InStr(pos + 5, result, "+")
        If plusPos = 0 Then plusPos = InStr(pos + 5, result, "+'")
        
        If plusPos > 0 Then
            Dim leftStart As Integer
            leftStart = pos + 5
            Do While leftStart < plusPos And Mid(result, leftStart, 1) = " "
                leftStart = leftStart + 1
            Loop
            Dim leftEnd As Integer
            leftEnd = plusPos - 1
            Do While leftEnd > leftStart And Mid(result, leftEnd, 1) = " "
                leftEnd = leftEnd - 1
            Loop
            Dim leftField As String
            leftField = Trim(Mid(result, leftStart, leftEnd - leftStart + 1))
            
            Dim quote1 As Integer
            quote1 = InStr(plusPos, result, "'")
            If quote1 > 0 Then
                Dim quote2 As Integer
                quote2 = InStr(quote1 + 1, result, "'")
                If quote2 > 0 Then
                    Dim separator As String
                    separator = Mid(result, quote1, quote2 - quote1 + 1)
                    
                    Dim rightPos As Integer
                    rightPos = InStr(quote2, result, "+")
                    If rightPos = 0 Then rightPos = InStr(quote2, result, "+'")
                    
                    If rightPos > 0 Then
                        Dim rightStart As Integer
                        rightStart = rightPos + 1
                        Do While rightStart <= Len(result) And Mid(result, rightStart, 1) = " "
                            rightStart = rightStart + 1
                        Loop
                        Dim rightEnd As Integer
                        rightEnd = rightStart
                        Do While rightEnd <= Len(result) And (Mid(result, rightEnd, 1) Like "[A-Za-z0-9_.]")
                            rightEnd = rightEnd + 1
                        Loop
                        Dim rightField As String
                        rightField = Trim(Mid(result, rightStart, rightEnd - rightStart))
                        
                        Dim newExpr As String
                        newExpr = "CONCAT(" & leftField & ", " & separator & ", ISNULL(" & rightField & ", ''))"
                        result = Left(result, leftStart - 1) & newExpr & Mid(result, rightEnd)
                    End If
                End If
            End If
        End If
        pos = InStr(pos + 1, result, "THEN ", vbTextCompare)
    Loop  ' <-- Cái này b?n có r?i
    
    FixSpecialSQL = result
End Function

Public Property Get TableDefs() As Collection
    ' B? check Nothing vì Class_Initialize dã t?o r?i
    If m_TableDefs.count = 0 Then
        LoadTableDefs
    End If
    Set TableDefs = m_TableDefs
End Property

Public Sub LoadTableDefs()
    Dim rs As Object
    Dim tdf As clsTableDef
    Dim SQL As String
    
    On Error GoTo Loi  ' Thêm dòng này d? b?t l?i
    
    Set m_TableDefs = New Collection
    
    ' Dùng tr?c ti?p, không qua OpenRecordset d? tránh ConvertSQL
    SQL = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQL, Conn, 3, 1  ' adOpenStatic, adLockReadOnly
    
    Do While Not rs.EOF
        Set tdf = New clsTableDef
        tdf.Init rs.Fields("TABLE_NAME").Value, Conn
        m_TableDefs.Add tdf, rs.Fields("TABLE_NAME").Value
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Exit Sub
    
Loi:
    Debug.Print "LoadTableDefs Error:", Err.number, Err.Description
End Sub

Public Property Get Relations() As Collection
    Set Relations = m_Relations
End Property

Public Function CreateRelation(Name As String, Table1 As String, Table2 As String, Optional Typ As Variant = 2) As Object
    Dim Rel As New clsRelation
    Rel.Init Name, Table1, Table2, 0, Conn
    m_Relations.Add Rel, Name
    Set CreateRelation = Rel
End Function

Public Sub NewPassword(OldPassword As String, NewPassword As String)
End Sub

Public Sub Closes()
    On Error Resume Next
    Conn.Close
End Sub

Public Function OpenDatabase(FileName As String, Exclusive As Boolean, ReadOnly As Boolean, Connect As String) As clsDatabase
    Dim db As New clsDatabase
    Dim sServer As String, sDB As String, sUser As String, sPwd As String
    sServer = "pc43\SQLEXPRESS"
    sDB = "thanhhuongbendinh"
    sUser = "sa"
    sPwd = "123456"
    db.Connect sServer, sDB, sUser, sPwd
    Set OpenDatabase = db
End Function

Public Sub BeginTrans()
    On Error GoTo Loi
    Conn.BeginTrans
    Exit Sub
Loi:
    Err.Raise Err.number, "clsDatabase.BeginTrans", Err.Description
End Sub

Public Sub CommitTrans()
    On Error GoTo Loi
    Conn.CommitTrans
    Exit Sub
Loi:
    Err.Raise Err.number, "clsDatabase.CommitTrans", Err.Description
End Sub

Public Sub RollbackTrans()
    On Error GoTo Loi
    Conn.RollbackTrans
    Exit Sub
Loi:
    Err.Raise Err.number, "clsDatabase.RollbackTrans", Err.Description
End Sub

Public Sub Rollback()
    On Error GoTo Loi
    Conn.RollbackTrans
    Exit Sub
Loi:
    Err.Raise Err.number, "clsDatabase.Rollback", Err.Description
End Sub

Public Function CreateTableDef(Name As String) As Object
    Dim tdf As New clsTableDef
    tdf.Init Name, Conn
    m_TableDefs.Add tdf, Name
    Set CreateTableDef = tdf
End Function

Public Function GetTableDef(TableName As String) As Object
    On Error Resume Next
    Dim tdf As Object
    If m_TableDefs.count = 0 Then
        LoadTableDefs
    End If
    Set tdf = m_TableDefs(TableName)
    If tdf Is Nothing Then
        Set tdf = CreateObject("clsTableDef")
        tdf.Init TableName, Conn
        m_TableDefs.Add tdf, TableName
    End If
    Set GetTableDef = tdf
End Function

Public Function CreateQueryDef(Name As String, SQL As String) As Object
    Dim qdf As New clsQueryDef
    On Error Resume Next
    m_QueryDefs.Remove Name
    On Error GoTo 0
    qdf.Init Name, SQL, Conn
    m_QueryDefs.Add qdf, Name
    Set CreateQueryDef = qdf
End Function

Public Property Get QueryDefByIndex(idx As Integer) As Object
    Dim i As Integer
    Dim item As Variant
    i = 1
    For Each item In m_QueryDefs
        If i = idx + 1 Then
            Set QueryDefByIndex = item
            Exit Property
        End If
        i = i + 1
    Next
    Set QueryDefByIndex = Nothing
End Property

Public Property Get QueryDefs() As Object
    Set QueryDefs = New clsQueryDefCollection
    QueryDefs.Init m_QueryDefs, Conn
End Property

Public Sub AddQuery(Name As String, SQL As String)
    On Error Resume Next
    m_QueryDefs.Remove Name
    On Error GoTo 0
    m_QueryDefs.Add SQL, Name
End Sub

Public Function GetLastRow(q As String, f As String) As Variant
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3
    rs.Open q, Conn, 3, 1
    If Not rs.EOF Then
        rs.MoveLast
        GetLastRow = rs.Fields(f).Value
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Sub CreateIndex(TableName As String, IndexName As String, ColumnName As String, Optional IsUnique As Boolean = False, Optional IsClustered As Boolean = False)
    Dim SQL As String
    Dim uniqueStr As String
    Dim clusteredStr As String
    SQL = "IF EXISTS (SELECT * FROM sys.indexes WHERE name = '" & IndexName & "' AND object_id = OBJECT_ID('" & TableName & "')) " & _
          "DROP INDEX " & IndexName & " ON " & TableName & "; "
    If IsUnique Then
        uniqueStr = "UNIQUE "
    End If
    If IsClustered Then
        clusteredStr = "CLUSTERED "
    End If
    SQL = SQL & "CREATE " & uniqueStr & clusteredStr & "INDEX " & IndexName & " ON " & TableName & " (" & ColumnName & ")"
    Me.ExecuteSQL SQL
End Sub
Public Function CreateView(ViewName As String, SQL As String) As Boolean
    On Error GoTo ErrHandler
    
    If Conn Is Nothing Or Conn.State <> 1 Then
        MsgBox "Chua k?t n?i!", vbCritical
        CreateView = False
        Exit Function
    End If
    
    ' ? S?A L?I ISNULL NGAY TRONG SQL
    Dim fixedSQL As String
    fixedSQL = SQL
    
    ' Thay IsNull(Field) thành Field IS NULL
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.pattern = "IsNull\s*\(\s*([^,]+?)\s*\)"
    fixedSQL = regex.Replace(fixedSQL, "$1 IS NULL")
    
    ' Chuy?n d?i SQL
    Dim strSQL As String
    strSQL = ConvertSQL(fixedSQL)
    
    ' Xóa ORDER BY
    Dim orderByPos As Integer
    orderByPos = InStr(1, strSQL, "ORDER BY", vbTextCompare)
    If orderByPos > 0 Then
        strSQL = Left(strSQL, orderByPos - 1)
    End If
    
    ' Xóa View cu
    On Error Resume Next
    Conn.Execute "DROP VIEW " & ViewName
    On Error GoTo ErrHandler
    
    ' T?o View m?i
    Conn.Execute "CREATE VIEW " & ViewName & " AS " & strSQL
    
    CreateView = True
    Exit Function
    
ErrHandler:
    MsgBox "L?i: " & Err.Description, vbCritical
    CreateView = False
End Function
