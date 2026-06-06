Attribute VB_Name = "clsIndex"
Option Explicit

Private m_Name As String
Private m_Fields As Collection
Private m_Unique As Boolean
Private m_Primary As Boolean
Private m_Conn As Object
Private m_TableName As String

Public Sub Init(TableName As String, Conn As Object)
    m_TableName = TableName
    Set m_Conn = Conn
    Set m_Fields = New Collection
    m_Unique = False
    m_Primary = False
    m_Name = ""  ' Kh?i t?o r?ng
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get Fields() As Collection
    Set Fields = m_Fields
End Property

Public Property Get Unique() As Boolean
    Unique = m_Unique
End Property

Public Property Let Unique(Value As Boolean)
    m_Unique = Value
End Property

Public Property Get Primary() As Boolean
    Primary = m_Primary
End Property

Public Property Let Primary(Value As Boolean)
    m_Primary = Value
End Property

Public Function CreateField(FieldName As String) As Object
    'Dim fld As New clsIndexField
    'fld.Name = FieldName
   ' m_Fields.Add fld, FieldName
    'Set CreateField = fld
End Function

Public Sub Append(Idx As Object)
    Dim sql As String
    Dim fld As Object
    Dim fieldList As String
    Dim i As Integer
    
    ' Xây d?ng danh sách field
    fieldList = ""
    For i = 1 To m_Fields.count
        fieldList = fieldList & "[" & m_Fields(i).Name & "]"
        If i < m_Fields.count Then fieldList = fieldList & ", "
    Next
    
    ' T?o câu l?nh
    If m_Primary Then
        sql = "ALTER TABLE [" & m_TableName & "] ADD CONSTRAINT [" & m_Name & "] PRIMARY KEY (" & fieldList & ")"
    Else
        sql = "CREATE "
        If m_Unique Then sql = sql & "UNIQUE "
        sql = sql & "INDEX [" & m_Name & "] ON [" & m_TableName & "] (" & fieldList & ")"
    End If
    
    On Error Resume Next
    m_Conn.Execute sql
    On Error GoTo 0
End Sub
