VERSION 5.00
Begin VB.Form frmSqlconnection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect sql server"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Info Connections"
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1200
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1800
         Width           =   3255
      End
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
         Top             =   3720
         Width           =   4455
      End
      Begin VB.TextBox Text4 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   3000
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   2400
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   3720
         TabIndex        =   5
         Text            =   "Truongthinh26"
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Text            =   "192.168.1.33,1433"
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nhãm"
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
         TabIndex        =   12
         Top             =   1200
         Width           =   735
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
         Top             =   3000
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
         Top             =   2400
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
         Top             =   1800
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
Dim isfirtload As Boolean

Private Sub Command1_Click()
    Screen.MousePointer = 11
    frmSplash.sqlIpServer = Text1.Text
    frmSplash.sqlDatabasename = Combo1.Text
    frmMain.sqlDatabasename = Combo1.Text
    frmSplash.sqlUsername = Text3.Text
    frmSplash.sqlPassword = Text4.Text
    Me.MousePointer = vbDefault
    isLogin = True

    Dim FileName As String
    Dim f As Integer
    Dim Line As String
    Dim newContent As String
    Dim DBName As String
    Dim UserName As String
    Dim Password As String

    DBName = Combo1.Text
    UserName = Text3.Text
    Password = Text4.Text

    FileName = App.path & "\Tools\Debug\SaovietTax.exe.config"

    If Dir(FileName) = "" Then
        MsgBox "Không tìm th?y file:" & vbCrLf & FileName, vbCritical
        Exit Sub
    End If

    ' ===== Ð?C FILE T?NG DÒNG =====
    f = FreeFile
    Open FileName For Input As #f
    newContent = ""
    
    Do While Not EOF(f)
        Line Input #f, Line
        
        ' N?u là dòng ch?a SqlConn (connection string thu?ng)
        If InStr(Line, "name=""SqlConn""") > 0 Then
            Line = "<add name=""SqlConn"" connectionString=""Server=192.168.1.33,1433;Database=" & DBName & ";User Id=" & UserName & ";Password=" & Password & ";"" providerName=""System.Data.SqlClient"" />"
        End If
        
        ' N?u là dòng ch?a DatablankEntities (Entity Framework)
        If InStr(Line, "name=""DatablankEntities""") > 0 Then
            Line = "<add name=""DatablankEntities"" connectionString=""metadata=res://*/Models.Model1.csdl|res://*/Models.Model1.ssdl|res://*/Models.Model1.msl;provider=System.Data.SqlClient;provider connection string=&quot;data source=192.168.1.33,1433;initial catalog=" & DBName & ";persist security info=True;user id=" & UserName & ";password=" & Password & ";trustservercertificate=True;MultipleActiveResultSets=True;App=EntityFramework&quot;"" providerName=""System.Data.EntityClient"" />"
        End If
        
        newContent = newContent & Line & vbCrLf
    Loop
    Close #f

    ' ===== GHI FILE =====
    f = FreeFile
    Open FileName For Output As #f
    Print #f, newContent;
    Close #f

    MsgBox "Ðã c?p nh?t config thành công!", vbInformation
    Unload Me
End Sub
Private Function ReplaceBetween(ByVal Source As String, ByVal StartText As String, ByVal EndText As String, ByVal NewValue As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim result As String
    
    result = Source
    startPos = InStr(1, result, StartText, vbTextCompare)
    
    If startPos > 0 Then
        endPos = InStr(startPos + Len(StartText), result, EndText, vbTextCompare)
        If endPos > 0 Then
            result = Left(result, startPos + Len(StartText) - 1) & NewValue & Mid(result, endPos)
        End If
    End If
    
    ReplaceBetween = result
End Function

Private Sub Combo2_Click()
    Dim cn As Object
    Dim rs As Object

    Set cn = CreateObject("ADODB.Connection")

    cn.Open "Provider=SQLOLEDB;" & _
            "Data Source=192.168.1.33,1433;" & _
            "Initial Catalog=SaovietAdmin;" & _
            "User ID=sa;" & _
            "Password=123456;"
    Dim strSQL As String

    strSQL = "SELECT d.*, c.PhancapName " & _
             "FROM tbDatabase d " & _
             "INNER JOIN tbCategory c ON " & _
           "    (c.Username = d.User1ID OR " & _
           "     c.Username = d.User2ID OR " & _
           "     c.Username = d.User3ID) " & _
             "WHERE c.PhancapName = '" & Combo2.Text & "'"

    Set rs = cn.Execute(strSQL)

    Set rs = cn.Execute(strSQL)
    Combo1.Clear

    Do Until rs.EOF
        Combo1.AddItem rs.Fields("Databasename").Value
        rs.MoveNext
    Loop
    If Combo1.ListCount > 0 Then
        Combo1.ListIndex = 0
    End If
    If isfirtload = True Then
        Dim i As Integer

        Dim phancnam As String
        Set cn = CreateObject("ADODB.Connection")

        cn.Open "Provider=SQLOLEDB;" & _
                "Data Source=192.168.1.33,1433;" & _
                "Initial Catalog=SaovietAdmin;" & _
                "User ID=sa;" & _
                "Password=123456;"

        Set rs = cn.Execute("SELECT * FROM tbCategory WHERE Username = '" & Text3.Text & "'")

        Do Until rs.EOF
            phancnam = rs.Fields("PhancapName").Value
            rs.MoveNext
        Loop

        For i = 0 To Combo2.ListCount - 1
            If Combo2.List(i) = phancnam Then
                Combo2.ListIndex = i
                Exit For
            End If
        Next i
        For i = 0 To Combo1.ListCount - 1
            If Combo1.List(i) = Text2.Text Then
                Combo1.ListIndex = i
                Exit For
            End If
        Next i
        isfirtload = False
    End If
End Sub
 

Private Sub Form_Activate()
    isLogin = False
End Sub

Private Sub Form_Load()
    isfirtload = True
    Dim configPath As String
    Dim fileContent As String
    Dim Lines() As String
    Dim Line As String
    Dim connString As String
    Dim server As String
    Dim database As String
    Dim userId As String
    Dim Password As String

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
        Password = GetValueFromConnString(connString, "Password")

        ' Gán vào TextBox
        Text1.Text = server
        Text2.Text = database
        Combo1.AddItem database
        Combo1.ListIndex = 0
        Text3.Text = userId
        Text4.Text = Password

        'MsgBox "Ðã l?y thông tin thành công!", vbInformation
    Else
        MsgBox "Không tìm th?y connection string!", vbExclamation
    End If

    Dim rs_tonkho As Object
    Dim Query As String
    Query = " SELECT name From sys.Databases Where database_id > 4 ORDER BY name"

    Dim cn As Object
    Dim rs As Object

    Set cn = CreateObject("ADODB.Connection")

    cn.Open "Provider=SQLOLEDB;" & _
            "Data Source=192.168.1.33,1433;" & _
            "Initial Catalog=SaovietAdmin;" & _
            "User ID=sa;" & _
            "Password=123456;"

    Set rs = cn.Execute("SELECT * FROM tbCategory")

    Do Until rs.EOF
        Combo2.AddItem rs.Fields("PhancapName").Value
        rs.MoveNext
    Loop
    If Combo2.ListCount > 0 Then
        Combo2.ListIndex = 0
    End If
    rs.Close
    cn.Close
    
End Sub

Private Function GetValueFromConnString(ByVal connString As String, ByVal key As String) As String
    Dim parts() As String
    Dim i As Integer
    Dim Item As String

    parts = Split(connString, ";")

    For i = 0 To UBound(parts)
        Item = Trim(parts(i))
        If InStr(Item, "=") > 0 Then
            If Trim(Split(Item, "=")(0)) = key Then
                GetValueFromConnString = Trim(Split(Item, "=")(1))
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

