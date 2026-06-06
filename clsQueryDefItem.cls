VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueryDefItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_SQL As String
Private m_Name As String
Private m_Conn As Object

Public Sub Init(SQL As Variant, Name As String, Conn As Object)
    If IsNull(SQL) Then
        m_SQL = ""
    ElseIf VarType(SQL) = vbString Then
        m_SQL = SQL
    Else
        m_SQL = ""
    End If
    m_Name = Name
    Set m_Conn = Conn
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get SQL() As String
    SQL = m_SQL
End Property

Public Property Let SQL(Value As String)
    m_SQL = Value
    
    On Error GoTo Loi
    ' Ki?m tra k?t n?i
    If m_Conn Is Nothing Then
        MsgBox "Chua k?t n?i SQL Server"
        Exit Property
    End If
    
    ' Chuy?n d?i SQL
    Dim s As String
    s = Value
    s = Replace(s, "DISTINCTROW", "DISTINCT")
    
    ' Xóa view cu
    On Error Resume Next
    m_Conn.Execute "IF OBJECT_ID('" & m_Name & "', 'V') IS NOT NULL DROP VIEW [" & m_Name & "]"
    On Error GoTo 0
    
    ' T?o view m?i
    Dim createSQL As String
    createSQL = "CREATE VIEW [" & m_Name & "] AS " & s
    m_Conn.Execute createSQL
    
    Exit Property
Loi:
    MsgBox "L?i t?o view: " & Err.Description
End Property
