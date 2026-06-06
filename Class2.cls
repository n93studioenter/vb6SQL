VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Name As String
Private m_SQL As String
Private m_Conn As Object

Public Sub Init(Name As String, SQL As String, Conn As Object)
    m_Name = Name
    m_SQL = SQL
    Set m_Conn = Conn
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get SQL() As String
    SQL = m_SQL
End Property

Public Property Let SQL(Value As String)
    m_SQL = Value
End Property

Public Function Execute(Optional Options As Variant = 0) As Object
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open m_SQL, m_Conn, 3, 3
    Set Execute = rs
End Function
