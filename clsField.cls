VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Name As String
Private m_Type As Long
Private m_Size As Long
Private m_IsPrimaryKey As Boolean
Private m_IsIdentity As Boolean
Private m_IsNullable As Boolean

Public Sub Init(Name As String, FieldType As Long, Optional FieldSize As Long = 0)
    m_Name = Name
    m_Type = FieldType
    m_Size = FieldSize
    m_IsPrimaryKey = False
    m_IsIdentity = False
    m_IsNullable = True
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get Typ() As Long
    Typ = m_Type
End Property

Public Property Let Typ(Value As Long)
    m_Type = Value
End Property

Public Property Get Size() As Long
    Size = m_Size
End Property

Public Property Let Size(Value As Long)
    m_Size = Value
End Property

Public Property Get IsPrimaryKey() As Boolean
    IsPrimaryKey = m_IsPrimaryKey
End Property

Public Property Let IsPrimaryKey(Value As Boolean)
    m_IsPrimaryKey = Value
End Property

Public Property Get IsIdentity() As Boolean
    IsIdentity = m_IsIdentity
End Property

Public Property Let IsIdentity(Value As Boolean)
    m_IsIdentity = Value
End Property

Public Property Get IsNullable() As Boolean
    IsNullable = m_IsNullable
End Property

Public Property Let IsNullable(Value As Boolean)
    m_IsNullable = Value
End Property

Public Sub Attributes(Attrs As Long)
    ' dbAutoIncrField = 16
    If (Attrs And 16) <> 0 Then
        m_IsIdentity = True
    End If
End Sub
