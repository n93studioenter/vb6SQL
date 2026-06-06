VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Name As String
Private m_Table1 As String
Private m_Table2 As String
Private m_Attributes As Long
Private m_Fields As Collection
Private m_Conn As Object

Public Sub Init(Name As String, Table1 As String, Table2 As String, Attributes As Long, Conn As Object)
    m_Name = Name
    m_Table1 = Table1
    m_Table2 = Table2
    m_Attributes = Attributes
    Set m_Conn = Conn
    Set m_Fields = New Collection
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Fields() As Collection
    Set Fields = m_Fields
End Property

Public Function CreateField(FieldName As String) As Object
    Dim fld As New clsRelationField
    fld.Name = FieldName
    m_Fields.Add fld, FieldName
    Set CreateField = fld
End Function

Public Sub Append(Rel As Object)
    Dim sql As String
    Dim fld As Object
    Dim fldLocal As String
    Dim fldForeign As String
    
    For Each fld In m_Fields
        fldLocal = fld.Name
        fldForeign = fld.ForeignName
        Exit For
    Next
    
    sql = "ALTER TABLE [" & m_Table2 & "] ADD CONSTRAINT [" & m_Name & "] FOREIGN KEY ([" & fldLocal & "]) REFERENCES [" & m_Table1 & "]([" & fldForeign & "])"
    
    On Error Resume Next
    m_Conn.Execute sql
    On Error GoTo 0
End Sub
