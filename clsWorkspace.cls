VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkspace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Trong clsWorkspace
Private m_DB As clsDatabase

Public Function OpenDatabase(FileName As String, Exclusive As Boolean, ReadOnly As Boolean, Optional Connect As String = "") As clsDatabase
    Set m_DB = New clsDatabase
     'm_DB.Connect "pc43\SQLEXPRESS", "thanhhuongbendinh", "sa", "123456"
    m_DB.Connect frmSplash.sqlIpServer, frmSplash.sqlDatabasename, frmSplash.sqlUsername, frmSplash.sqlPassword

    Set OpenDatabase = m_DB
End Function

Public Sub BeginTrans()
    If Not m_DB Is Nothing Then
        m_DB.BeginTrans   ' G?i t? clsDatabase
    End If
End Sub

Public Sub CommitTrans()
    If Not m_DB Is Nothing Then
        m_DB.CommitTrans
    End If
End Sub

Public Sub RollbackTrans()
    If Not m_DB Is Nothing Then
        m_DB.RollbackTrans
    End If
End Sub
Public Sub Rollback()
    If Not m_DB Is Nothing Then
        m_DB.Rollback
    End If
End Sub
' Ðóng workspace
Public Sub Closes()
    If Not m_DB Is Nothing Then
        m_DB.Closes
        Set m_DB = Nothing
    End If
End Sub
