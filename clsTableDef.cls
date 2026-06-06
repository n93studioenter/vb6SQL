VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTableDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Name As String
Private m_Fields As Collection
Private m_Indexes As Collection
Private m_Conn As Object

Public Sub Init(Name As String, Conn As Object)
    m_Name = Name
    Set m_Conn = Conn
    Set m_Fields = New Collection
    Set m_Indexes = New Collection
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Fields() As Collection
    Set Fields = m_Fields
End Property

Public Property Get Indexes() As Collection
    Set Indexes = m_Indexes
End Property

Public Function CreateField(FieldName As String, Optional FieldType As Long = 10, Optional FieldSize As Long = 255) As Object
    Dim fld As New clsField
    fld.Init FieldName, FieldType, FieldSize
    m_Fields.Add fld, FieldName
    Set CreateField = fld
End Function

Public Function CreateIndex(IndexName As String) As Object
   ' Dim Idx As New clsIndex
    'Idx.Init m_Name, m_Conn  ' C?n truy?n thêm tham s?
    'Idx.Name = IndexName
   ' m_Indexes.Add Idx, IndexName
   ' Set CreateIndex = Idx
End Function

Public Sub Append(tdf As Object)
    Dim sql As String
    Dim fld As Object
    Dim fieldDefs As String
    
    ' Xây d?ng câu l?nh CREATE TABLE
    sql = "CREATE TABLE [" & m_Name & "] ("
    
    For Each fld In m_Fields
        fieldDefs = fieldDefs & "[" & fld.Name & "] " & GetSQLType(fld.Type, fld.Size)
        
        If fld.IsPrimaryKey Then
            fieldDefs = fieldDefs & " PRIMARY KEY"
        End If
        
        If fld.IsIdentity Then
            fieldDefs = fieldDefs & " IDENTITY(1,1)"
        End If
        
        If Not fld.IsNullable Then
            fieldDefs = fieldDefs & " NOT NULL"
        End If
        
        fieldDefs = fieldDefs & ", "
    Next
    
    ' Xóa d?u ph?y cu?i
    If Right(fieldDefs, 2) = ", " Then
        fieldDefs = Left(fieldDefs, Len(fieldDefs) - 2)
    End If
    
    sql = sql & fieldDefs & ")"
    
    ' Th?c thi t?o b?ng
    On Error Resume Next
    m_Conn.Execute sql
    If Err.number <> 0 Then
        Err.Raise Err.number, "clsTableDef.Append", Err.Description
    End If
    On Error GoTo 0
End Sub

Private Function GetSQLType(FieldType As Long, FieldSize As Long) As String
    Select Case FieldType
        Case 1:      GetSQLType = "BIT"                ' dbBoolean
        Case 2:      GetSQLType = "TINYINT"            ' dbByte
        Case 3:      GetSQLType = "SMALLINT"           ' dbInteger
        Case 4:      GetSQLType = "INT"                ' dbLong
        Case 5:      GetSQLType = "MONEY"              ' dbCurrency
        Case 6:      GetSQLType = "REAL"               ' dbSingle
        Case 7:      GetSQLType = "FLOAT"              ' dbDouble
        Case 8:      GetSQLType = "DATETIME"           ' dbDate
        Case 10:     GetSQLType = "NVARCHAR(" & FieldSize & ")"  ' dbText
        Case 11:     GetSQLType = "VARBINARY(MAX)"     ' dbLongBinary
        Case 12:     GetSQLType = "NVARCHAR(MAX)"      ' dbMemo
        Case Else:   GetSQLType = "NVARCHAR(255)"
    End Select
End Function
