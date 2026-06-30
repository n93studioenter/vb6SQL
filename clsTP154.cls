VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls154"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Kieu vat tu
Public MaSo As Long             ' Ma vat tu
Public MaPhanLoai As Long       ' ma phan loai
Public sohieu As String         ' So hieu vat tu
Public TenVattu As String       ' Ten vat tu
Public DonVi As String          ' Don vi tinh
Public GhiChu As String
Public MaTK As Long
'======================================================================================
' Thu tuc khoi tao object tu ma so vat tu
' mvt:  Ma so vat tu can tham chieu
'       =0 neu vat tu moi
'======================================================================================
Public Sub InitTPMaSo(mvt As Long)
    Dim rs_vattu As Object
    
    If mvt > 0 Then
        Set rs_vattu = DBKetoan.OpenRecordset("SELECT * FROM TP154 WHERE MaSo=" + CStr(mvt), dbOpenSnapshot)
        If rs_vattu.recordCount = 0 Then GoTo iv
        MaSo = rs_vattu!MaSo
        MaPhanLoai = rs_vattu!MaPhanLoai
        sohieu = rs_vattu!sohieu
        TenVattu = rs_vattu!TenVattu
        DonVi = rs_vattu!DonVi
        GhiChu = rs_vattu!GhiChu
        MaTK = rs_vattu!MaTK
        rs_vattu.Close
        Set rs_vattu = Nothing
    Else
iv:
        MaSo = 0
        MaPhanLoai = 0
        sohieu = ""
        TenVattu = ""
        DonVi = ""
        GhiChu = "..."
        MaTK = 0
    End If
End Sub
'==========================================================================
' Thu tuc khoi tao object tu so hieu vat tu
' mvt:  So hieu vat tu can tham chieu
'==========================================================================
Public Function InitTPSohieu(shvt As String) As String
    Dim SQL As String
    
    SQL = "SELECT DISTINCTROW MaSo AS F1 FROM TP154 WHERE Sohieu='" + shvt + "'"
    
    InitTPMaSo SelectSQL(SQL)
    InitTPSohieu = TenVattu
End Function
'==========================================================================
' Thu tuc dang ky vat tu
'==========================================================================
Public Function GhiTP() As Integer
    Dim SQL As String
        
    SQL = "INSERT INTO TP154 (MaSo,MaPhanLoai,SoHieu,TenVattu,DonVi,GhiChu,MaTK) VALUES (" _
        + CStr(Lng_MaxValue("MaSo", "TP154") + 1) + "," + CStr(MaPhanLoai) + ",'" + sohieu + "','" + TenVattu + "','" + DonVi + "','" + GhiChu + "'," + CStr(MaTK) + ")"
    If ExecuteSQL5(SQL) <> 0 Then
        GhiTP = -1
    Else
        GhiTP = 0
        MaSo = Lng_MaxValue("MaSo", "TP154")
    End If
End Function
'==========================================================================
' Cap nhat thong tin ve vat tu
'==========================================================================
Public Function SuaTP() As Integer
    Dim SQL As String
    SQL = "UPDATE TP154 SET MaPhanLoai=" + CStr(MaPhanLoai) + ",SoHieu='" + sohieu + "',TenVattu='" + TenVattu + "',DonVi='" + DonVi + "',GhiChu='" + GhiChu + "',MaTK=" + CStr(MaTK) + " WHERE MaSo=" + CStr(MaSo)
    If ExecuteSQL5(SQL) <> 0 Then
        SuaTP = -1
    Else
        SuaTP = 0
    End If
End Function
'==========================================================================
' Thu tuc xoa mot vat tu
'==========================================================================
Public Function XoaTP() As Integer
    Dim TK As New ClsTaikhoan
    XoaTP = -1
    If SelectSQL("SELECT MaSo AS F1 FROM ChungTu WHERE MaTP=" + CStr(MaSo)) <> 0 Then Exit Function
    If ExecuteSQL5("DELETE FROM TP154 WHERE MaSo=" + CStr(MaSo)) = 0 Then XoaTP = 0
    TK.InitTaikhoanMaSo MaTK
    TK.NoDauKy = SelectSQL("SELECT SUM(DK) AS F1 FROM TP154 WHERE MaTK=" + CStr(MaTK))
    TK.CoDauKy = 0
    TK.CapNhatTk
    
    Set TK = Nothing
End Function
Public Function Kiem_tra_qd15() As Integer
 Dim KT As Integer
 Dim SQL  As String
'    Dim rs_tk As Object
'    kt = 0
'     Set rs_tk = DBKetoan.OpenRecordset("select * from HethongTK where SoHieu LIKE '621%'", dbOpenSnapshot)
'     Do While Not rs_tk.EOF
'       kt = 1
'     Loop
      SQL = "SELECT SUM(maso) AS F1 from HethongTK where SoHieu LIKE '621%'"
      Kiem_tra_qd15 = SelectSQL(SQL)
End Function
Public Function SoCPNVL(tdau As Integer, tcuoi As Integer) As Double
    Dim SQL As String
   
   If Kiem_tra_qd15() > 0 Then
   
                  SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE HethongTK.SoHieu LIKE '621%' AND " _
        + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
    SoCPNVL = SelectSQL(SQL)
    SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) + " WHERE HethongTK.SoHieu LIKE '152%' AND TK.SoHieu LIKE '621%' AND " _
        + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
    SoCPNVL = SoCPNVL - SelectSQL(SQL)
    If SoCPNVL = 0 Then
        SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) + " WHERE HethongTK.SoHieu LIKE '154%' AND TK.SoHieu LIKE '621%' AND " _
            + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
        SoCPNVL = SelectSQL(SQL)
    End If
Else
    If SoCPNVL = 0 Then
        SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) + " WHERE HethongTK.SoHieu LIKE '154%' and( TK.SoHieu not LIKE '214%' and TK.SoHieu not LIKE '334%') AND " _
         + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
        SoCPNVL = SelectSQL(SQL)
         
         SQL = " SELECT SUM(SoPS) AS F1 FROM (ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) WHERE HethongTK.SoHieu LIKE '154%' and ChungTu.MaTKCo = 0 and " _
        + WThang("ThangCT", tdau, tcuoi) + " And chungtu.MaTP = " + CStr(MaSo)
             SoCPNVL = SoCPNVL + SelectSQL(SQL)
    End If
    End If
End Function

Public Function SoCPNC(tdau As Integer, tcuoi As Integer) As Double
    Dim SQL As String
    If Kiem_tra_qd15() > 0 Then
     SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE HethongTK.SoHieu LIKE '622%' AND " _
        + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
    SoCPNC = SelectSQL(SQL)
    If SoCPNC = 0 Then
        SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) + " WHERE HethongTK.SoHieu LIKE '154%' AND TK.SoHieu LIKE '622%' AND " _
            + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
        SoCPNC = SelectSQL(SQL)
    End If
Else
SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE HethongTK.SoHieu LIKE '334%' AND " _
        + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
    SoCPNC = SelectSQL(SQL)
    If SoCPNC = 0 Then
        SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) + " WHERE HethongTK.SoHieu LIKE '154%' AND TK.SoHieu LIKE '334%' AND " _
            + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
        SoCPNC = SelectSQL(SQL)
        SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) + " WHERE HethongTK.SoHieu LIKE '334%' AND TK.SoHieu LIKE '154%' AND " _
            + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
        
        SoCPNC = SoCPNC - SelectSQL(SQL)
    End If

    End If
End Function

Public Function SoCPKHTT(tdau As Integer, tcuoi As Integer) As Double
    Dim SQL As String
    
    SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE MaLoai=12 AND TK_ID2=" + CStr(TKCPSX_ID) + " AND " _
        + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
    SoCPKHTT = SelectSQL(SQL)
End Function

Public Function SoSanLuong(tdau As Integer, tcuoi As Integer) As Double
    Dim SQL As String
    
    SQL = "SELECT SUM(SoLuong) AS F1 FROM ThanhPham WHERE " + WThang("Thang", tdau, tcuoi) + " AND Ma154=" + CStr(MaSo)
    SoSanLuong = SelectSQL(SQL)
End Function

Public Function SoCPM(tdau As Integer, tcuoi As Integer) As Double
    Dim SQL As String
    If Kiem_tra_qd15() > 0 Then
            SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE HethongTK.SoHieu LIKE '623%' AND " _
                + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
            SoCPM = SelectSQL(SQL)
            If SoCPM = 0 Then
                SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) + " WHERE HethongTK.SoHieu LIKE '154%' AND TK.SoHieu LIKE '623%' AND " _
                    + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
                SoCPM = SelectSQL(SQL)
            End If
Else

    SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE HethongTK.SoHieu LIKE '214%' AND " _
        + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
    SoCPM = SelectSQL(SQL)
    If SoCPM = 0 Then
        SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) + " WHERE HethongTK.SoHieu LIKE '154%' AND TK.SoHieu LIKE '214%' AND " _
            + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
        SoCPM = SelectSQL(SQL)
    End If
End If
    
End Function

Public Function SoCPNVLPB(tdau As Integer, tcuoi As Integer) As Double
    Dim SQL As String, i As Integer
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        SQL = SQL + "+CPNVLPB" + CStr(i)
    Next
    SQL = "SELECT (" + SQL + ") AS F1 FROM TP154 WHERE MaSo=" + CStr(MaSo)
    SoCPNVLPB = SelectSQL(SQL)
End Function

Public Function SoCPNCPB(tdau As Integer, tcuoi As Integer) As Double
    Dim SQL As String, i As Integer
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        SQL = SQL + "+CPNCPB" + CStr(i)
    Next
    SQL = "SELECT (" + SQL + ") AS F1 FROM TP154 WHERE MaSo=" + CStr(MaSo)
    SoCPNCPB = SelectSQL(SQL)
End Function

Public Function SoCPMPB(tdau As Integer, tcuoi As Integer) As Double
    Dim SQL As String, i As Integer
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        SQL = SQL + "+CPMPB" + CStr(i)
    Next
    SQL = "SELECT (" + SQL + ") AS F1 FROM TP154 WHERE MaSo=" + CStr(MaSo)
    SoCPMPB = SelectSQL(SQL)
End Function

Public Function SoCPKH(tdau As Integer, tcuoi As Integer) As Double
    Dim SQL As String, i As Integer
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        SQL = SQL + "+CPKH_" + CStr(i)
    Next
    SQL = "SELECT (" + SQL + ") AS F1 FROM TP154 WHERE MaSo=" + CStr(MaSo)
    SoCPKH = SelectSQL(SQL)
End Function

Public Function SoCPSXC(tdau As Integer, tcuoi As Integer) As Double
    Dim SQL As String, i As Integer
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        SQL = SQL + "+CPSXC_" + CStr(i)
    Next
    SQL = "SELECT (" + SQL + ") AS F1 FROM TP154 WHERE MaSo=" + CStr(MaSo)
    SoCPSXC = SelectSQL(SQL)
End Function

Public Function SoCPSXCTT(tdau As Integer, tcuoi As Integer) As Double
    Dim SQL As String
    
    SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE HethongTK.SoHieu LIKE '627%' AND " _
        + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
    SoCPSXCTT = SelectSQL(SQL)
    If SoCPSXCTT = 0 Then
        SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) + " WHERE HethongTK.SoHieu LIKE '154%' AND TK.SoHieu LIKE '627%' AND " _
            + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
        SoCPSXCTT = SelectSQL(SQL)
    End If
End Function

Public Function SoCPBH(tdau As Integer, tcuoi As Integer) As Double
    Dim SQL As String
    
    SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE HethongTK.SoHieu LIKE '641%' AND " _
        + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
    SoCPBH = SelectSQL(SQL)
End Function

Public Function SoCPQL(tdau As Integer, tcuoi As Integer) As Double
    Dim SQL As String, kq As Double
    
    SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) + " WHERE (HethongTK.SoHieu LIKE '911%' AND (TK.SoHieu LIKE '142%' OR TK.SoHieu LIKE '242%')) AND " _
        + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
    kq = SelectSQL(SQL)
    
    SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE HethongTK.SoHieu LIKE '642%' AND " _
        + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
    SoCPQL = SelectSQL(SQL) + kq
End Function

Public Function SoDT(tdau As Integer, tcuoi As Integer) As Double
    Dim SQL As String
    
    SQL = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(1) + " WHERE HethongTK.TK_ID=" + CStr(TKDT_ID) + " AND " _
        + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(MaSo)
    SoDT = SelectSQL(SQL)
    If SoDT = 0 Then
        SQL = "SELECT SUM(SoPS) AS F1 FROM (" + ChungTu2TKNC(1) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo WHERE HethongTK.TK_ID=" + CStr(TKDT_ID) + " AND " _
            + WThang("ThangCT", tdau, tcuoi) + " AND Vattu.SoHieu='" + CStr(sohieu) + "'"
        SoDT = SelectSQL(SQL)
    End If
End Function

Public Function SoDK() As Double
    Dim SQL As String
    
    SQL = "SELECT DK AS F1 FROM TP154 WHERE MaSo=" + CStr(MaSo)
    SoDK = SelectSQL(SQL)
End Function

Public Sub GhiDK(dk As Double)
    Dim TK As New ClsTaikhoan
    
    ExecuteSQL5 "UPDATE TP154 SET DK=" + DoiDau(dk) + " WHERE MaSo=" + CStr(MaSo)
    TK.InitTaikhoanMaSo MaTK
    TK.NoDauKy = SelectSQL("SELECT SUM(DK) AS F1 FROM TP154 WHERE MaTK=" + CStr(MaTK))
    TK.CoDauKy = 0
    TK.CapNhatTk
    
    Set TK = Nothing
End Sub

Public Function GiaThanhCK(thang As Integer) As Double
    Dim X As Double
        
    If thang > 0 Then
        GiaThanhCK = SoDK() + SoCPNVL(pThangDauKy, thang) + SoCPNC(pThangDauKy, thang) + SoCPM(pThangDauKy, thang) + SoCPNVLPB(pThangDauKy, thang) + SoCPNCPB(pThangDauKy, thang) + SoCPMPB(pThangDauKy, thang) + SoCPSXCTT(pThangDauKy, thang) + SoCPSXC(pThangDauKy, thang)
        X = SelectSQL("SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(1) + " WHERE HethongTK.SoHieu LIKE '154%' AND MaTP=" + CStr(MaSo) + " AND " + WThang("ThangCT", 0, thang))
        'If GiaThanhCK >= x Then
            GiaThanhCK = GiaThanhCK - X
        'Else
        '    GiaThanhCK = GiaThanhCK - SelectSQL("SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(1) + " WHERE HethongTK.SoHieu LIKE '154%' AND MaTP=" + CStr(MaSo) + " AND " + WThang("ThangCT", thang, thang))
        'End If
    Else
        GiaThanhCK = SoDK()
    End If
End Function

Public Function ChiPhiTP(thang As Integer) As Double
    If thang > 0 Then
        ChiPhiTP = SoDK() + SoCPNVL(pThangDauKy, thang) + SoCPNC(pThangDauKy, thang) + SoCPM(pThangDauKy, thang) + SoCPNVLPB(pThangDauKy, thang) + SoCPNCPB(pThangDauKy, thang) + SoCPMPB(pThangDauKy, thang) + SoCPSXCTT(pThangDauKy, thang) + SoCPSXC(pThangDauKy, thang)
    Else
        ChiPhiTP = SoDK()
    End If
End Function

Public Sub GhiCPTT(tdau As Integer, tcuoi As Integer)
    Dim i As Integer
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        ExecuteSQL5 "UPDATE TP154 SET CPBHTT" + CStr(i) + "=" + DoiDau(SoCPBH(i, i)) + ",CPQLTT" + CStr(i) + "=" + DoiDau(SoCPQL(i, i)) + ",CPSXCTT" + CStr(i) + "=" + DoiDau(SoCPSXCTT(i, i)) + " WHERE MaSo=" + CStr(MaSo)                        ' ,CPKH_" + CStr(i) + "=" + DoiDau(SoCPKHTT(i, i)) + "
    Next
    ExecuteSQL5 "UPDATE TP154 SET CPNVL=" + DoiDau(SoCPNVL(tdau, tcuoi)) + ",CPNC=" + DoiDau(SoCPNC(tdau, tcuoi)) + ",CPM=" + DoiDau(SoCPM(tdau, tcuoi)) + ",DT=" + DoiDau(SoDT(tdau, tcuoi)) + ",SanLuong=" + DoiDau(SoSanLuong(tdau, tcuoi)) + " WHERE MaSo=" + CStr(MaSo)
End Sub

Public Sub GhiCPPB(tdau As Integer, tcuoi As Integer, cpkh As Double, cpsxc As Double)
    Dim i As Integer, SQL As String
    
    For i = CThangDB(tdau) To CThangDB(ThangTruoc(tcuoi))
        SQL = SQL + ",CPKH_" + CStr(i) + "=0,CPSXC_" + CStr(i) + "=0"
    Next
    ExecuteSQL5 "UPDATE TP154 SET CPNC=CPNC" + SQL + " WHERE MaSo=" + CStr(MaSo)
    SQL = CStr(CThangDB(tcuoi))
    ExecuteSQL5 "UPDATE TP154 SET CPKH_" + SQL + "=" + DoiDau(cpkh) + ",CPSXC_" + SQL + "=" + DoiDau(cpsxc) + " WHERE MaSo=" + CStr(MaSo)
End Sub

Public Sub GhiCPPB2(tdau As Integer, tcuoi As Integer, cpbh As Double, cpql As Double, cptc As Double)
    Dim i As Integer, SQL As String
    
    For i = CThangDB(tdau) To CThangDB(ThangTruoc(tcuoi))
        SQL = SQL + ",CPBH" + CStr(i) + "=0,CPQL" + CStr(i) + "=0,CPTC" + CStr(i) + "=0"
    Next
    ExecuteSQL5 "UPDATE TP154 SET CPNC=CPNC" + SQL + " WHERE MaSo=" + CStr(MaSo)
    SQL = CStr(CThangDB(tcuoi))
    ExecuteSQL5 "UPDATE TP154 SET CPBH" + SQL + "=" + DoiDau(cpbh) + ",CPQL" + SQL + "=" + DoiDau(cpql) + ",CPTC" + SQL + "=" + DoiDau(cptc) + " WHERE MaSo=" + CStr(MaSo)
End Sub

Public Sub GhiCPPB3(tdau As Integer, tcuoi As Integer, cpnc As Double, cpm As Double)
    Dim i As Integer, SQL As String
    
    For i = CThangDB(tdau) To CThangDB(ThangTruoc(tcuoi))
        SQL = SQL + ",CPNCPB" + CStr(i) + "=0,CPMPB" + CStr(i) + "=0"
    Next
    ExecuteSQL5 "UPDATE TP154 SET CPNC=CPNC" + SQL + " WHERE MaSo=" + CStr(MaSo)
    SQL = CStr(CThangDB(tcuoi))
    ExecuteSQL5 "UPDATE TP154 SET CPNCPB" + SQL + "=" + DoiDau(cpnc) + ",CPMPB" + SQL + "=" + DoiDau(cpm) + " WHERE MaSo=" + CStr(MaSo)
End Sub

Public Sub XDDauKy(thang As Integer)
    ExecuteSQL5 "UPDATE TP154 SET DK1=" + DoiDau(GiaThanhCK(ThangTruoc(thang))) + " WHERE MaSo=" + CStr(MaSo)
End Sub

Public Sub XDCuoiKy(thang As Integer)
    ExecuteSQL5 "UPDATE TP154 SET CK1=" + DoiDau(GiaThanhCK(thang)) + " WHERE MaSo=" + CStr(MaSo)
End Sub

Public Sub DanXuat(tp As Cls154)
    With tp
        .MaSo = MaSo
        .MaPhanLoai = MaPhanLoai
        .sohieu = sohieu
        .TenVattu = TenVattu
        .DonVi = DonVi
        .GhiChu = GhiChu
        .MaTK = MaTK
    End With
End Sub

