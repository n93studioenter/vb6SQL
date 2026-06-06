VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsNhanVien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MaSo As Long
Public MaPhanLoai As Long
Public sohieu As String
Public Ten As String
Public DiaChi As String
Public NgaySinh As Date
Public CMND As String
Public Tel1 As String
Public Tel2 As String
Public email As String
Public taikhoan As String
Public GhiChu As String

Private Sub Class_Initialize()
    MaPhanLoai = 0
    sohieu = "..."
    Ten = "..."
    DiaChi = "..."
    CMND = "..."
    Tel1 = "..."
    Tel2 = "..."
    email = "..."
    taikhoan = "..."
    GhiChu = "..."
End Sub

Public Function GhiNV() As Integer
    Dim kq As Integer
    kq = ExecuteSQL5("INSERT INTO NhanVien (MaSo, MaPhanLoai, SoHieu, Ten, DiaChi, NgaySinh, CMND, Tel1, Tel2, EMail, TaiKhoan, GhiChu) " _
        & " VALUES (" + CStr(Lng_MaxValue("MaSo", "NhanVien") + 1) + "," + CStr(MaPhanLoai) + ",'" + sohieu + "','" + Ten + "','" + DiaChi + "',#" + Format(NgaySinh, Mask_DB) + "#,'" + CMND + "','" + Tel1 + "','" + Tel2 _
        + "','" + email + "','" + taikhoan + "','" + GhiChu + "')")
    If kq = 0 Then MaSo = Lng_MaxValue("MaSo", "NhanVien")
    GhiNV = kq
End Function

Public Function SuaNV() As Integer
    SuaNV = ExecuteSQL5("UPDATE NhanVien SET MaPhanLoai=" + CStr(MaPhanLoai) + ",SoHieu='" + sohieu + "',Ten='" + Ten + "',DiaChi='" + DiaChi _
        + "',NgaySinh=#" + Format(NgaySinh, Mask_DB) + "#,CMND='" + CMND + "',Tel1='" + Tel1 + "',Tel2='" + Tel2 _
        + "',EMail='" + email + "',TaiKhoan='" + taikhoan + "',GhiChu='" + GhiChu + "' WHERE MaSo=" + CStr(MaSo))
End Function

Public Function XoaNV() As Integer
    If SelectSQL("SELECT TOP 1 MaNV AS F1 FROM ChungTu WHERE MaNV>0") = 0 Then
        XoaNV = ExecuteSQL5("DELETE * FROM NhanVien WHERE MaSo=" + CStr(MaSo))
    Else
        XoaNV = -1
    End If
End Function

Public Sub InitNVMaSo(ms As Long)
    Dim rs As Object
    
    If ms > 0 Then
        Set rs = DBKetoan.OpenRecordset("SELECT * FROM NhanVien WHERE MaSo=" + CStr(ms), dbOpenSnapshot)
        If rs.recordCount = 0 Then GoTo Init
        On Error Resume Next
        MaSo = rs!MaSo
        MaPhanLoai = rs!MaPhanLoai
        sohieu = rs!sohieu
        Ten = rs!Ten
        DiaChi = rs!DiaChi
        NgaySinh = rs!NgaySinh
        CMND = rs!CMND
        Tel1 = rs!Tel1
        Tel2 = rs!Tel2
        email = rs!email
        taikhoan = rs!taikhoan
        GhiChu = rs!GhiChu
        On Error GoTo 0
        rs.Close
        Set rs = Nothing
    Else
Init:
        Class_Initialize
    End If

End Sub

Public Function InitNVSoHieu(sh As String) As Long
    Dim ms As Long
    
    ms = SelectSQL("SELECT MaSo AS F1 FROM NhanVien WHERE SoHieu='" + sh + "'")
    If ms > 0 Then InitNVMaSo ms
    InitNVSoHieu = ms
End Function
