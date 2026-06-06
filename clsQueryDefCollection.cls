VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueryDefCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_QueryDefs As Collection
Private m_Conn As Object

Public Sub Init(QueryDefs As Collection, Conn As Object)
    Set m_QueryDefs = QueryDefs
    Set m_Conn = Conn
End Sub

Public Function Item(Index As Variant) As Object
    On Error GoTo Loi
    Dim result As New clsQueryDefItem
    Dim sqlValue As Variant
    
    ' L?y giá tr? t? collection
    sqlValue = m_QueryDefs(Index)
    
    result.Init sqlValue, CStr(Index), m_Conn
    Set Item = result
    Exit Function
    
Loi:
    ' N?u không tìm th?y, t?o m?i
    Dim newItem As New clsQueryDefItem
    newItem.Init "", CStr(Index), m_Conn
    Set Item = newItem
End Function
