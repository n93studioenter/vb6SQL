VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsIndex - Class Module
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
    m_Name = ""
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

' Cách 1: Dùng CreateField (không c?n clsIndexField)
Public Function CreateField(fieldName As String) As Object
    m_Fields.Add fieldName, fieldName
    Set CreateField = Nothing
End Function

' Cách 2: Dùng AddField (khuy?n ngh?)
Public Sub AddField(fieldName As String)
    m_Fields.Add fieldName, fieldName
End Sub

Public Sub Append()
    Dim SQL As String
    Dim fieldList As String
    Dim i As Integer
    
    If m_Fields.count = 0 Then
        Debug.Print "No fields defined for index"
        Exit Sub
    End If
    
    ' Xây d?ng danh sách field
    fieldList = ""
    For i = 1 To m_Fields.count
        fieldList = fieldList & "[" & m_Fields(i) & "]"
        If i < m_Fields.count Then fieldList = fieldList & ", "
    Next
    
    ' T?o câu l?nh SQL
    If m_Primary Then
        SQL = "ALTER TABLE [" & m_TableName & "] ADD CONSTRAINT [" & m_Name & "] PRIMARY KEY (" & fieldList & ")"
    Else
        SQL = "CREATE "
        If m_Unique Then SQL = SQL & "UNIQUE "
        SQL = SQL & "INDEX [" & m_Name & "] ON [" & m_TableName & "] (" & fieldList & ")"
    End If
    
    On Error Resume Next
    m_Conn.Execute SQL
    If Err.number <> 0 Then
        Debug.Print "L?i t?o index: " & Err.Description
    Else
        Debug.Print "T?o index thành công: " & m_Name
    End If
    On Error GoTo 0
End Sub

