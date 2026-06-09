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
    If m_Indexes.count = 0 Then LoadIndexes
    Set Indexes = m_Indexes
End Property

Private Sub LoadIndexes()
    Dim rs As Object
    Dim idx As clsIndex
    Dim SQL As String
    
    Set m_Indexes = New Collection
    
    SQL = "SELECT i.name, i.is_unique, i.is_primary_key, " & _
          "COL_NAME(ic.object_id, ic.column_id) AS ColName " & _
          "FROM sys.indexes i " & _
          "INNER JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id " & _
          "WHERE i.object_id = OBJECT_ID('" & m_Name & "') " & _
          "AND i.is_hypothetical = 0 AND i.type <> 0 " & _
          "ORDER BY i.name, ic.key_ordinal"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQL, m_Conn, 3, 1
    
    Dim lastIdx As String
    lastIdx = ""
    
    Do While Not rs.EOF
        If rs!Name <> lastIdx Then
            Set idx = New clsIndex
            idx.Init m_Name, m_Conn
            idx.Name = rs!Name
            idx.Unique = rs!is_unique
            idx.Primary = rs!is_primary_key
            idx.AddField rs!ColName
            m_Indexes.Add idx, rs!Name
            lastIdx = rs!Name
        Else
            m_Indexes(lastIdx).AddField rs!ColName
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Public Function CreateField(FieldName As String, Optional FieldType As Long = 10, Optional FieldSize As Long = 255) As Object
    Dim fld As New clsField
    fld.Init FieldName, FieldType, FieldSize
    m_Fields.Add fld, FieldName
    Set CreateField = fld
End Function

Public Function CreateIndex(IndexName As String) As Object
    Dim idx As New clsIndex
    idx.Init m_Name, m_Conn
    idx.Name = IndexName
    m_Indexes.Add idx, IndexName
    Set CreateIndex = idx
End Function

Public Sub Append(tdf As Object)
    Dim SQL As String
    Dim fld As Object
    Dim fieldDefs As String
    
    SQL = "CREATE TABLE [" & m_Name & "] ("
    
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
    
    If Right(fieldDefs, 2) = ", " Then
        fieldDefs = Left(fieldDefs, Len(fieldDefs) - 2)
    End If
    
    SQL = SQL & fieldDefs & ")"
    
    On Error Resume Next
    m_Conn.Execute SQL
    If Err.number <> 0 Then
        Err.Raise Err.number, "clsTableDef.Append", Err.Description
    End If
    On Error GoTo 0
End Sub

Private Function GetSQLType(FieldType As Long, FieldSize As Long) As String
    Select Case FieldType
        Case 1:      GetSQLType = "BIT"
        Case 2:      GetSQLType = "TINYINT"
        Case 3:      GetSQLType = "SMALLINT"
        Case 4:      GetSQLType = "INT"
        Case 5:      GetSQLType = "MONEY"
        Case 6:      GetSQLType = "REAL"
        Case 7:      GetSQLType = "FLOAT"
        Case 8:      GetSQLType = "DATETIME"
        Case 10:     GetSQLType = "NVARCHAR(" & FieldSize & ")"
        Case 11:     GetSQLType = "VARBINARY(MAX)"
        Case 12:     GetSQLType = "NVARCHAR(MAX)"
        Case Else:   GetSQLType = "NVARCHAR(255)"
    End Select
End Function

