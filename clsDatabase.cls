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

Public Function Connect(Server As String, db As String, uID As String, PWD As String) As Boolean
    On Error Resume Next
    Conn.Open "Provider=SQLOLEDB;Data Source=" & Server & ";Initial Catalog=" & db & ";User ID=" & uID & ";Password=" & PWD & ";"
    Connect = (Err.number = 0)
End Function

Public Sub ExecuteSQL(SQL As String, Optional Options As Variant = dbFailOnError)
    Debug.Print "Mientru " & SQL
    On Error GoTo Loi
    ' ADO m?c d?nh dã có co ch? báo l?i nhu dbFailOnError
    Conn.Execute SQL, RecordsAffected
    Exit Sub
Loi:
    'Err.Raise Err.number, "clsDatabase.ExecuteSQL", Err.Description
    Debug.Print Err.number, "clsDatabase.ExecuteSQL", Err.Description
End Sub
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
        
        ' B? qua n?i dung trong chu?i (gi?a 2 d?u ')
        If ch = "'" Then
            inString = Not inString
        End If
        
        ' Tìm d?u #
        If Not inString And ch = "#" Then
            startPos = i
            i = i + 1
            dateStr = ""
            
            ' Ð?c n?i dung gi?a 2 d?u #
            Do While i <= Len(result)
                ch = Mid(result, i, 1)
                If ch = "#" Then
                    Exit Do
                End If
                dateStr = dateStr & ch
                i = i + 1
            Loop
            
            ' Chuy?n d?i d?nh d?ng ngày tháng
            dateStr = ConvertDateFormat(dateStr)
            
            ' Thay th? b?ng d?nh d?ng SQL Server
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
    
    ' X? lý các d?nh d?ng ngày tháng
    dateStr = Trim(dateStr)
    
    ' Ð?nh d?ng mm/dd/yyyy ho?c mm/dd/yy
    If InStr(dateStr, "/") > 0 Then
        parts = Split(dateStr, "/")
        If UBound(parts) = 2 Then
            month = parts(0)
            day = parts(1)
            year = parts(2)
            
            ' Chuy?n d?i nam 2 s? thành 4 s?
            If Len(year) = 2 Then
                If CInt(year) >= 30 Then
                    year = "19" & year
                Else
                    year = "20" & year
                End If
            End If
            
            ' Chu?n hóa thành YYYY-MM-DD
            ConvertDateFormat = year & "-" & Right("0" & month, 2) & "-" & Right("0" & day, 2)
            Exit Function
        End If
    End If
    
    ' Ð?nh d?ng dd-mm-yyyy
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
    
    ' Không xác d?nh du?c d?nh d?ng, tr? v? nguyên b?n
    ConvertDateFormat = dateStr
End Function
Private Function ConvertSQL(ByVal SQL As String) As String
    Dim result As String
    result = SQL

    ' S?a CONVERT tru?c
    result = FixConvertError(result)

    ' S?a n?i chu?i d?ng field + 'text' + field
    result = FixPlusConcat(result)

    ' S?a ELSE '' END+field
    result = FixElseEndConcat(result)

    ' Các rule khác...
    result = Replace(result, "DISTINCTROW", "DISTINCT")
    result = Replace(result, "distinctrow", "DISTINCT")

    result = ConvertDateDelimiters(result)
    result = ConvertWildcards(result)
    result = ConvertIIFToCase(result)

    result = Replace(result, "Format(", "CONVERT(varchar, ")
    result = FixFormatDate(result)  ' Thêm hàm x? lý Format ngày tháng
    
    result = Replace(result, "Fix(", "ROUND(")
    result = Replace(result, "Date()", "GETDATE()")
    result = Replace(result, "Now()", "GETDATE()")

    ConvertSQL = result
End Function

' Thêm hàm này d? x? lý Format ngày tháng
Private Function FixFormatDate(ByVal SQL As String) As String
    Dim result As String
    result = SQL
    
    ' X? lý Format(date, "dd/mm/yy") -> CONVERT(varchar, date, 3)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    
    ' Pattern: Format(tên_field, 'd?nh_d?ng')
    regex.pattern = "Format\(([^,]+),\s*['""]([^'""]+)['""]\)"
    
    Dim matches As Object
    Set matches = regex.Execute(result)
    
    Dim match As Object
    Dim FieldName As String
    Dim dateFormat As String
    Dim sqlCode As String
    
    For Each match In matches
        FieldName = match.SubMatches(0)
        dateFormat = LCase(match.SubMatches(1))
        
        ' Chuy?n d?i d?nh d?ng Access sang SQL Server
        Select Case dateFormat
            Case "dd/mm/yy", "dd/mm/yyyy", "dd\/mm\/yy", "dd\/mm\/yyyy"
                sqlCode = "CONVERT(varchar, " & FieldName & ", 3)"
            Case "mm/dd/yy", "mm/dd/yyyy"
                sqlCode = "CONVERT(varchar, " & FieldName & ", 1)"
            Case "yyyymmdd"
                sqlCode = "CONVERT(varchar, " & FieldName & ", 112)"
            Case "dd/mm/yyyy hh:nn:ss", "dd/mm/yy hh:nn:ss"
                sqlCode = "CONVERT(varchar, " & FieldName & ", 103)"
            Case Else
                sqlCode = "CONVERT(varchar, " & FieldName & ", 103)"  ' Default dd/mm/yyyy
        End Select
        
        result = Replace(result, match.Value, sqlCode)
    Next
    
    FixFormatDate = result
End Function

' S?a l?i CONVERT tham s? 3
Private Function FixConvertError(ByVal SQL As String) As String
    Dim result As String
    result = SQL

    ' X? lý CONVERT(varchar, field, 'dd/mm/yy') -> CONVERT(varchar, field, 3)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.pattern = "CONVERT\s*\(\s*varchar\s*,\s*([^,]+)\s*,\s*['""]([^'""]+)['""]\s*\)"

    Dim matches As Object
    Set matches = regex.Execute(result)

    Dim match As Object
    Dim FieldName As String
    Dim dateFormat As String
    Dim formatCode As String

    For Each match In matches
        FieldName = match.SubMatches(0)
        dateFormat = LCase(match.SubMatches(1))

        ' Chuy?n d?i d?nh d?ng sang s? code
        Select Case dateFormat
        Case "dd/mm/yy", "dd/mm/yyyy", "dd\/mm\/yy", "dd\/mm\/yyyy"
            formatCode = "3"
        Case "mm/dd/yy", "mm/dd/yyyy"
            formatCode = "1"
        Case "yyyymmdd"
            formatCode = "112"
        Case Else
            formatCode = "3"  ' Default
        End Select

        result = Replace(result, match.Value, "CONVERT(varchar, " & FieldName & ", " & formatCode & ")")
    Next

    FixConvertError = result
End Function
Private Function FixThenConcat(ByVal SQL As String) As String
    Dim result As String
    result = SQL
    
    ' Tìm pattern: THEN field + 'text' + field
    ' X? lý c? tru?ng h?p: THEN field+' - ' + field (không có kho?ng tr?ng)
    
    ' Pattern 1: THEN field + 'text' + field (có kho?ng tr?ng)
    Dim pos As Integer
    pos = InStr(1, result, "THEN ", vbTextCompare)
    
    Do While pos > 0
        ' Tìm d?u + (có ho?c không có kho?ng tr?ng)
        Dim plusPos As Integer
        plusPos = InStr(pos + 5, result, "+")
        
        ' N?u không tìm th?y, th? tìm "+'" (không có kho?ng tr?ng)
        If plusPos = 0 Then
            plusPos = InStr(pos + 5, result, "+'")
        End If
        
        If plusPos > 0 Then
            ' L?y field bên trái
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
            
            ' L?y separator
            Dim quote1 As Integer
            quote1 = InStr(plusPos, result, "'")
            If quote1 > 0 Then
                Dim quote2 As Integer
                quote2 = InStr(quote1 + 1, result, "'")
                If quote2 > 0 Then
                    Dim separator As String
                    separator = Mid(result, quote1, quote2 - quote1 + 1)
                    
                    ' Tìm field bên ph?i
                    Dim rightPos As Integer
                    rightPos = InStr(quote2, result, "+")
                    If rightPos = 0 Then
                        rightPos = InStr(quote2, result, "+'")
                    End If
                    
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
                        
                        ' T?o bi?u th?c m?i
                        Dim newExpr As String
                        newExpr = "CONCAT(" & leftField & ", " & separator & ", ISNULL(" & rightField & ", ''))"
                        
                        ' Thay th?
                        result = Left(result, leftStart - 1) & newExpr & Mid(result, rightEnd)
                    End If
                End If
            End If
        End If
        
        pos = InStr(pos + 1, result, "THEN ", vbTextCompare)
    Loop
    
    FixThenConcat = result
End Function
Private Function FixElseEndConcat(ByVal SQL As String) As String
    Dim result As String
    result = SQL
    
    ' Tìm pattern: ELSE '' END+field
    Dim pos As Integer
    pos = InStr(1, result, "ELSE '' END+", vbTextCompare)
    
    Do While pos > 0
        ' L?y tên field sau d?u +
        Dim startField As Integer
        startField = pos + 11  ' Sau "ELSE '' END+"
        
        Dim endField As Integer
        endField = startField
        Do While endField <= Len(result) And (Mid(result, endField, 1) Like "[A-Za-z0-9_.]")
            endField = endField + 1
        Loop
        
        Dim FieldName As String
        FieldName = Mid(result, startField, endField - startField)
        
        ' Thay th?
        Dim newText As String
        newText = "ELSE ISNULL(" & FieldName & ", '') END"
        
        result = Left(result, pos - 1) & newText & Mid(result, endField)
        
        pos = InStr(pos + 1, result, "ELSE '' END+", vbTextCompare)
    Loop
    
    FixElseEndConcat = result
End Function
' S?a CONVERT(varchar, field,'dd/mm/yy') -> CONVERT(varchar, field, 103)
Private Function FixConvertDateFormat(ByVal SQL As String) As String
    Dim result As String
    result = SQL
    
    ' Tìm pattern: CONVERT(varchar, field,'dd/mm/yy')
    ' Thay b?ng: CONVERT(varchar, field, 103)
    
    Dim pos As Integer
    pos = InStr(1, result, "CONVERT(varchar,", vbTextCompare)
    
    Do While pos > 0
        ' Tìm d?u ph?y d?u tiên
        Dim comma1 As Integer
        comma1 = InStr(pos + 16, result, ",")
        
        If comma1 > 0 Then
            ' Tìm d?u ph?y th? hai
            Dim comma2 As Integer
            comma2 = InStr(comma1 + 1, result, ",")
            
            If comma2 > 0 Then
                ' Tìm d?u ngo?c dóng
                Dim endPos As Integer
                endPos = InStr(comma2 + 1, result, ")")
                
                If endPos > 0 Then
                    ' L?y field name
                    Dim FieldName As String
                    FieldName = Trim(Mid(result, pos + 16, comma1 - pos - 16))
                    
                    ' L?y format string
                    Dim formatStr As String
                    formatStr = Trim(Mid(result, comma1 + 1, comma2 - comma1 - 1))
                    formatStr = Replace(formatStr, "'", "")
                    formatStr = Replace(formatStr, """", "")
                    
                    ' Xác d?nh mã format SQL Server
                    Dim formatCode As String
                    Select Case formatStr
                        Case "dd/mm/yy", "dd/mm/yyyy", "dd\/mm\/yyyy"
                            formatCode = "103"
                        Case "mm/dd/yy", "mm/dd/yyyy"
                            formatCode = "101"
                        Case "yyyymmdd"
                            formatCode = "112"
                        Case Else
                            formatCode = "103"  ' default
                    End Select
                    
                    ' T?o câu l?nh m?i
                    Dim newConvert As String
                    newConvert = "CONVERT(varchar, " & FieldName & ", " & formatCode & ")"
                    
                    ' Thay th?
                    result = Left(result, pos - 1) & newConvert & Mid(result, endPos + 1)
                End If
            End If
        End If
        
        pos = InStr(pos + 1, result, "CONVERT(varchar,", vbTextCompare)
    Loop
    
    FixConvertDateFormat = result
End Function
Private Function FixCaseWhenConcat(ByVal SQL As String) As String
    Dim result As String
    result = SQL
    
    ' Tìm pattern: ELSE '' END+DienGiai -> ELSE ISNULL(DienGiai, '') END
    Dim pos As Integer
    pos = InStr(1, result, "ELSE '' END+", vbTextCompare)
    
    Do While pos > 0
        Dim FieldName As String
        Dim startField As Integer
        startField = pos + 11  ' Sau "ELSE '' END+"
        
        Dim endField As Integer
        endField = startField
        Do While endField <= Len(result) And (Mid(result, endField, 1) Like "[A-Za-z0-9_.]")
            endField = endField + 1
        Loop
        
        FieldName = Mid(result, startField, endField - startField)
        
        Dim newText As String
        newText = "ELSE ISNULL(" & FieldName & ", '') END"
        
        result = Left(result, pos - 1) & newText & Mid(result, endField)
        
        pos = InStr(pos + 1, result, "ELSE '' END+", vbTextCompare)
    Loop
    
    FixCaseWhenConcat = result
End Function
 
' Hàm chuy?n d?i Chr() thành CHAR()
Private Function ConvertChrFunction(ByVal SQL As String) As String
    Dim result As String
    Dim pos As Integer
    Dim startPos As Integer
    Dim endPos As Integer
    Dim numValue As String

    result = SQL
    pos = InStr(1, result, "Chr(", vbTextCompare)

    Do While pos > 0
        startPos = pos
        endPos = pos + 4

        ' Tìm v? trí dóng ngo?c
        Do While endPos <= Len(result)
            If Mid(result, endPos, 1) = ")" Then
                Exit Do
            End If
            endPos = endPos + 1
        Loop

        ' L?y s? bên trong
        numValue = Mid(result, pos + 4, endPos - pos - 4)

        ' Thay th? Chr(9) thành CHAR(9)
        result = Left(result, startPos - 1) & "CHAR(" & numValue & ")" & Mid(result, endPos + 1)

        pos = InStr(1, result, "Chr(", vbTextCompare)
    Loop

    ConvertChrFunction = result
End Function

' Hàm chuy?n d?i IIF thành CASE WHEN
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

        ' Tìm v? trí k?t thúc c?a IIF(...)
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

        ' Tìm d?u ph?y th? nh?t
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

        ' Tìm d?u ph?y th? hai
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

        ' B? qua n?i dung trong chu?i
        If ch = "'" Then
            inString = Not inString
        End If

        ' Phát hi?n t? LIKE (không n?m trong chu?i)
        If Not inString Then
            If i + 4 <= Len(result) Then
                If UCase(Mid(result, i, 5)) = " LIKE" Then
                    inLike = True
                    i = i + 5
                    ' Không dùng Continue Do, ch? tang i và ti?p t?c vòng l?p
                    GoTo NextIteration
                End If
            End If
        End If

        ' X? lý pattern c?a LIKE
        If inLike And Not inString And ch = "'" Then
            patternStart = i + 1
            pattern = ""
            newPattern = ""
            i = i + 1

            ' Ð?c pattern
            Do While i <= Len(result)
                ch = Mid(result, i, 1)
                If ch = "'" Then
                    Exit Do
                End If
                pattern = pattern & ch
                i = i + 1
            Loop

            ' Chuy?n d?i wildcard
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

            ' Thay th? pattern m?i
            result = Left(result, patternStart - 1) & newPattern & Mid(result, i)
            inLike = False
        End If

NextIteration:
        i = i + 1
    Loop

    ConvertWildcards = result
End Function

' Hàm chuy?n d?i n?i chu?i
' Hàm chuy?n d?i n?i chu?i - Phiên b?n don gi?n
Private Function ConvertStringConcat(ByVal SQL As String) As String
    Dim result As String
    result = SQL
    
    ' Chuy?n d?i field + 'text' + field -> CONCAT(field, 'text', field)
    Dim pos As Integer
    pos = InStr(1, result, "+ '", vbTextCompare)
    
    If pos > 0 Then
        ' L?y field bên trái
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
        
        ' L?y separator
        Dim quote1 As Integer, quote2 As Integer
        quote1 = InStr(pos, result, "'")
        quote2 = InStr(quote1 + 1, result, "'")
        Dim separator As String
        separator = Mid(result, quote1, quote2 - quote1 + 1)
        
        ' L?y field bên ph?i
        Dim rightPos As Integer
        rightPos = InStr(quote2, result, "+")
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
        
        ' T?o CONCAT
        Dim newExpr As String
        newExpr = "CONCAT(" & leftField & ", " & separator & ", ISNULL(" & rightField & ", ''))"
        
        ' Thay th?
        result = Replace(result, leftField & " + " & separator & " + " & rightField, newExpr)
    End If
    
    ConvertStringConcat = result
End Function

Private Function IsLetterOrDigit(ByVal ch As String) As Boolean
    IsLetterOrDigit = (ch >= "A" And ch <= "Z") Or _
                      (ch >= "a" And ch <= "z") Or _
                      (ch >= "0" And ch <= "9") Or _
                      ch = "_"
End Function
' OpenRecordset - 2 tham s? (dúng v?i code cu)
' Hàm s?a các câu CONCAT b? l?i cú pháp
' S?a CONCAT b? l?i cú pháp
Private Function FixBrokenConcat(ByVal SQL As String) As String
    Dim result As String
    result = SQL
    
    ' S?a CONCAT(, -> CONCAT('',
    Do While InStr(result, "CONCAT(,") > 0
        result = Replace(result, "CONCAT(,", "CONCAT('',")
    Loop
    
    ' S?a )CONCAT( -> ) + CONCAT(
    Do While InStr(result, ")CONCAT(") > 0
        result = Replace(result, ")CONCAT(", ") + CONCAT(")
    Loop
    
    ' S?a ''))CONCAT( -> '')) + CONCAT(
    Do While InStr(result, "''))CONCAT(") > 0
        result = Replace(result, "''))CONCAT(", "'')) + CONCAT(")
    Loop
    
    FixBrokenConcat = result
End Function
' ========== HÀM X? LÝ Ð?C BI?T CHO CÂU SQL C?A B?N ==========
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
                    Dim FieldName As String
                    FieldName = Trim(Mid(result, pos + 16, comma1 - pos - 16))
                    Dim newConvert As String
                    newConvert = "CONVERT(varchar, " & FieldName & ", 103)"
                    result = Left(result, pos - 1) & newConvert & Mid(result, endPos + 1)
                End If
            End If
        End If
        pos = InStr(pos + 1, result, "CONVERT(varchar,", vbTextCompare)
    Loop
    
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
    Loop
    
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
    Loop
    
    FixSpecialSQL = result
End Function
Public Function OpenRecordset(Source As String, Optional Typ As Variant = 2, Optional Typ2 As Variant = 2) As Object
    Dim rs As Object
    Dim SQL As String

    SQL = Source
    ' ========== G?I HÀM X? LÝ Ð?C BI?T ==========
    SQL = FixSpecialSQL(SQL)
    SQL = ConvertSQL(SQL)

    ' Lo?i b? DISTINCTROW - SQL Server không h? tr?
    SQL = Replace(SQL, "DISTINCTROW", "DISTINCT")
    SQL = Replace(SQL, "distinctrow", "DISTINCT")
    SQL = Trim(SQL)  ' Xóa kho?ng tr?ng th?a
    ' Ki?m tra QueryDef
    On Error Resume Next
    SQL = m_QueryDefs(Source)

    On Error GoTo 0

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3  ' adUseClient
    rs.CursorType = 3      ' adOpenStatic (cho phép MoveLast)
    rs.LockType = 1        ' adLockReadOnly
    ' Chuy?n d?i Type
    Dim rsType As Integer
    Select Case Typ
    Case 2:    ' dbOpenSnapshot
        rsType = 3   ' adOpenStatic
    Case Else:
        rsType = 3
    End Select
    Debug.Print SQL
    rs.Open SQL, Conn, rsType, 3
    Set OpenRecordset = rs
End Function
' Hàm chuy?n d?i n?i chu?i d?ng field + 'text' + field
Private Function FixPlusConcat(ByVal SQL As String) As String
    Dim result As String
    result = SQL
    
    ' Tìm pattern: field + 'text' + field
    Dim pattern As String
    Dim replacement As String
    
    ' Ví d?: SoHieu + ' - ' + TenPhanLoai -> CONCAT(SoHieu, ' - ', ISNULL(TenPhanLoai, ''))
    Dim pos As Integer
    pos = InStr(1, result, "+ '", vbTextCompare)
    
    If pos > 0 Then
        ' Tìm field bên trái
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
        
        ' L?y separator (chu?i gi?a 2 d?u nháy)
        Dim quote1 As Integer
        quote1 = InStr(pos, result, "'")
        Dim quote2 As Integer
        quote2 = InStr(quote1 + 1, result, "'")
        Dim separator As String
        separator = Mid(result, quote1, quote2 - quote1 + 1)
        
        ' Tìm field bên ph?i
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
            
            ' T?o CONCAT
            Dim newConcat As String
            newConcat = "CONCAT(" & leftField & ", " & separator & ", ISNULL(" & rightField & ", ''))"
            
            ' Thay th?
            result = Replace(result, leftField & " + " & separator & " + " & rightField, newConcat)
        End If
    End If
    
    FixPlusConcat = result
End Function
' ========== HÀM X? LÝ Ð?C BI?T CHO CÂU SQL ==========
' ========== HÀM S?A L?I CONVERT ==========
 

 
' Thêm property này vào clsDatabase
Public Property Get QueryDefByIndex(Idx As Integer) As Object
    Dim i As Integer
    Dim Item As Variant
    i = 1
    For Each Item In m_QueryDefs
        If i = Idx + 1 Then  ' Vì Collection b?t d?u t? 1
            Set QueryDefByIndex = Item
            Exit Property
        End If
        i = i + 1
    Next
    Set QueryDefByIndex = Nothing
End Property

' Property QueryDefs tr? v? collection h? tr? gán .sql
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

' TableDefs
Public Property Get TableDefs() As Collection
    Set TableDefs = m_TableDefs
End Property

' Relations
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
    ' SQL Server không h? tr?, b? qua
End Sub

Public Sub Closes()
    On Error Resume Next
    Conn.Close
End Sub

Public Function OpenDatabase(FileName As String, Exclusive As Boolean, ReadOnly As Boolean, Connect As String) As clsDatabase
    Dim db As New clsDatabase

    ' Parse connection string t? Connect d? l?y thông tin SQL Server
    ' Ho?c d?c t? Registry/file c?u hình
    Dim sServer As String, sDB As String, sUser As String, sPwd As String

    ' L?y t? Registry (có th? set s?n)
    'sServer = GetSetting(App.EXEName, "SQLServer", "Server", "localhost")
    'sDB = GetSetting(App.EXEName, "SQLServer", "Database", "KeToan")
    'sUser = GetSetting(App.EXEName, "SQLServer", "User", "sa")
    'sPwd = GetSetting(App.EXEName, "SQLServer", "Password", "")
    sServer = "pc43\SQLEXPRESS"
    sDB = "thanhhuongbendinh"
    sUser = "sa"
    sPwd = "123456"



    db.Connect sServer, sDB, sUser, sPwd
    Set OpenDatabase = db
End Function
' Trong clsDatabase
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
' Rollback transaction
Public Sub Rollback()
    On Error GoTo Loi
    Conn.RollbackTrans
    Exit Sub
Loi:
    Err.Raise Err.number, "clsDatabase.Rollback", Err.Description
End Sub
' T?o TableDef
Public Function CreateTableDef(Name As String) As Object
    Dim tdf As New clsTableDef
    tdf.Init Name, Conn
    m_TableDefs.Add tdf, Name
    Set CreateTableDef = tdf
End Function
 
' Thêm hàm này vào clsDatabase
Public Sub LoadTableDefs()
    Dim rs As Object
    Dim tdf As Object
    Dim SQL As String

    ' Xóa collection cu
    Set m_TableDefs = New Collection

    ' L?y danh sách b?ng t? SQL Server
    SQL = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME"
    Set rs = Me.OpenRecordset(SQL, dbOpenSnapshot)

    Do While Not rs.EOF
        Set tdf = New clsTableDef
        tdf.Init rs.Fields("TABLE_NAME").Value, Conn
        m_TableDefs.Add tdf, rs.Fields("TABLE_NAME").Value
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub
' T?o QueryDef
Public Function CreateQueryDef(Name As String, SQL As String) As Object
    Dim qdf As New clsQueryDef

    On Error Resume Next
    m_QueryDefs.Remove Name
    On Error GoTo 0

    qdf.Init Name, SQL, Conn

    m_QueryDefs.Add qdf, Name

    Set CreateQueryDef = qdf
End Function
Public Function GetLastRow(q As String, f As String) As Variant
    Dim rs As Object
    
    ' T?o recordset v?i cursor client d? h? tr? MoveLast
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3  ' adUseClient (3)
    rs.Open q, Conn, 3, 1  ' adOpenStatic (3), adLockReadOnly (1)
    
    If Not rs.EOF Then
        rs.MoveLast
        GetLastRow = rs.Fields(f).Value
    End If
    
    rs.Close
    Set rs = Nothing
End Function
