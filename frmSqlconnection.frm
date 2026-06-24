VERSION 5.00
Begin VB.Form frmSqlconnection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect sql server"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Info Connections"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton Command1 
         Caption         =   "KÕt nèi"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   3240
         Width           =   4455
      End
      Begin VB.TextBox Text4 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   9
         Text            =   "123456"
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Text            =   "sa"
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   1440
         TabIndex        =   5
         Text            =   "Truongthinh26"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Text            =   "pc43\SQLEXPRESS"
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CSDL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "M¸y chñ"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Label Label2 
      Caption         =   "M¸y chñ"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
End
Attribute VB_Name = "frmSqlconnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isLogin As Boolean

Private Sub Command1_Click()
    Screen.MousePointer = 11
    frmSplash.sqlIpServer = Text1.Text
    frmSplash.sqlDatabasename = Text2.Text
    frmMain.sqlDatabasename = Text2.Text
    frmSplash.sqlUsername = Text3.Text
    frmSplash.sqlPassword = Text4.Text
    Me.MousePointer = vbDefault
    isLogin = True
    Unload Me
End Sub

Private Sub Form_Activate()
    isLogin = False
End Sub

Private Sub Form_Load()
    Dim configPath As String
    Dim fileContent As String
    Dim lines() As String
    Dim i As Integer
    Dim line As String
    Dim connString As String
    Dim server As String
    Dim database As String
    Dim userId As String
    Dim password As String
    
    ' Ðu?ng d?n d?n file config
    configPath = App.path & "\Tools\Debug\SaovietTax.exe.Config"
    
    ' Ki?m tra file t?n t?i
    If Dir(configPath) = "" Then
        MsgBox "Không tìm th?y file config!", vbExclamation
        Exit Sub
    End If
    
    ' Ð?c toàn b? file
    Open configPath For Input As #1
    fileContent = Input$(LOF(1), 1)
    Close #1
    
    ' Tìm connection string
    Dim startPos As Integer
    Dim endPos As Integer
    
    startPos = InStr(fileContent, "connectionString=""")
    If startPos > 0 Then
        startPos = startPos + Len("connectionString=""")
        endPos = InStr(startPos, fileContent, """")
        connString = Mid(fileContent, startPos, endPos - startPos)
        
        ' L?y thông tin
        server = GetValueFromConnString(connString, "Server")
        database = GetValueFromConnString(connString, "Database")
        userId = GetValueFromConnString(connString, "User Id")
        password = GetValueFromConnString(connString, "Password")
        
        ' Gán vào TextBox
        Text1.Text = server
        Text2.Text = database
        Text3.Text = userId
        Text4.Text = password
        
        'MsgBox "Ðã l?y thông tin thành công!", vbInformation
    Else
        MsgBox "Không tìm th?y connection string!", vbExclamation
    End If
End Sub

Private Function GetValueFromConnString(ByVal connString As String, ByVal key As String) As String
    Dim parts() As String
    Dim i As Integer
    Dim item As String

    parts = Split(connString, ";")

    For i = 0 To UBound(parts)
        item = Trim(parts(i))
        If InStr(item, "=") > 0 Then
            If Trim(Split(item, "=")(0)) = key Then
                GetValueFromConnString = Trim(Split(item, "=")(1))
                Exit Function
            End If
        End If
    Next i

    GetValueFromConnString = ""
End Function

Private Sub Form_Unload(Cancel As Integer)
    If isLogin = False Then
        End
    End If
End Sub

