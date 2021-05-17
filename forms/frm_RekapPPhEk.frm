VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_RekapPPhEk 
   ClientHeight    =   4545
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   12300
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   12300
   Begin VB.ListBox List1 
      Height          =   1530
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   12015
   End
   Begin VB.Frame Frame1 
      Caption         =   " Filter "
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   12015
      Begin VB.TextBox txt_Bulan 
         Height          =   315
         Left            =   3000
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txt_Tahun 
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cb_pph 
         Height          =   330
         Left            =   5280
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   1380
         Width           =   4335
      End
      Begin VB.ComboBox cb_divisi 
         Height          =   330
         Left            =   480
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   600
         Width           =   3135
      End
      Begin VB.ComboBox cb_kpp 
         Height          =   330
         Left            =   5280
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton cmd_load 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Load"
         Height          =   375
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "4. Jenis PPh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5040
         TabIndex        =   12
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "1. Unit / Divisi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "3. KPP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5040
         TabIndex        =   7
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bulan"
         Height          =   210
         Left            =   2400
         TabIndex        =   6
         Top             =   1372
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   480
         TabIndex        =   5
         Top             =   1372
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2. data s/d "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   870
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4290
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10557
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10557
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lb_caption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rekap PPh - Ekualisasi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   481
      Left            =   0
      TabIndex        =   0
      Top             =   117
      Width           =   12285
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   13215
   End
   Begin VB.Menu mncek 
      Caption         =   "cek"
   End
End
Attribute VB_Name = "frm_RekapPPhEk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nama_data As String


Sub disable_Form()
    'Me.Frame3.Enabled = False
    Me.Frame1.Enabled = False
End Sub

Sub Enable_Form()
    'Me.Frame3.Enabled = True
    Me.Frame1.Enabled = True
End Sub

Sub pph42_sheet_01(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim nmFile As String, sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    'Dim fLs As New Excel.Worksheet
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long
    Dim Total1 As Currency, totalPBKonstruksi As Currency, totalPBSewa As Currency
    Dim data1, data2, isi
    
    nmFile = App.Path & "\rep\uReport_KonsepRekapSPT4(2).xlsx"
    f = nmFile
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet01 - error open xls template")
            Exit Sub
        End If
    Else
        MsgBox "File template tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet01 - error open xls template")
        Exit Sub
    End If
    
    'open sheet 1, isi data looping
    Set fLs = fl.Sheets(1)
    baris = 2
    kolom = 2
    
    fLs.Cells(baris, kolom) = "REKAP SSP + PPH " & Me.cb_pph & " - Tahun " & tahun
    baris = 4
    fLs.Cells(baris, kolom) = "Filter Unit: " & DIVISI
    baris = 5
    fLs.Cells(baris, kolom) = "Filter Bulan: " & Masa_Pajak
    baris = 6
    fLs.Cells(baris, kolom) = "Filter KPP: " & kpp
    'fLs.Cells(baris, kolom + 3) = 2500000
    'fLs.Cells(baris, kolom + 3).NumberFormat = "#,##0"
                                    
    baris = 7
    kolom = 5
    fLs.Cells(baris, kolom) = "Data dari awal s/d Bulan:" & Masa_Pajak & " tahun " & _
                            tahun
    
    '-- get data per KPP
    sql = "Select mkpp.npwp, mkpp.kpp_administrasi, " & _
        "F_get_pph42_sewa_bruto('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp,'') as PB_Sewa," & _
        "F_get_ssp_pph('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp,'','PPH FINAL SEWA') as SSP_Sewa," & _
        "'' as p1," & _
        "F_get_pph42_konstruksi_bruto('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp,'') as PB_Konstruksi," & _
        "F_get_ssp_pph('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp,'','PPH FINAL') as SSP_Konstruksi," & _
        "'' as p2," & _
        "F_get_pph42_obligasi_bruto('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp,'') as PB_Bunga_Obligasi," & _
        "'' as SSP_Bunga_Obligasi, '0' as p3," & _
        "'' as PB_Deviden, '' as SSP_Deviden, '' as p4," & _
        "F_get_pph42_sewa_bruto('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp,'') + " & _
        "F_get_pph42_konstruksi_bruto('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp,'') +" & _
        "F_get_pph42_obligasi_bruto('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp,'') as Total_PB, " & _
        "F_get_ssp_pph('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp,'','PPH FINAL SEWA') + " & _
        "F_get_ssp_pph('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp,'','PPH FINAL') as Total_SSP, " & _
        "'' as p5 " & _
        "From mkpp "
    If Trim(UCase(Me.cb_kpp.text)) = "ALL" Then
        sql = sql & "order by npwp"
    Else
        sql = sql & "where npwp = '" & get_kode_combo(Me.cb_kpp, "#") & "' " & _
        "order by npwp"
    End If
        
    
    Call setListInfo(Me.List1, "Proses Sheet01 - Open Query..")
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql errir", "", sql)
        Exit Sub
    End If
    
    Call setListInfo(Me.List1, "Proses Sheet01 - Open Query..OK")
    jRec = RecordCount(rs)
    
    totalPBKonstruksi = 0
    totalPBSewa = 0
    baris = 9
    If jRec > 0 Then
        rs.MoveFirst
        c = 1
        Do While rs.EOF = False
            Call info_progress(Me.StatusBar1, 1, c, jRec, "write data perKPP")
            fLs.Cells(baris, 3) = "'" & cek_null(rs(0))
            fLs.Cells(baris, 4) = cek_null(rs(1))
            
            fLs.Cells(baris, 5) = cek_Money(rs(2))
            totalPBSewa = totalPBSewa + cek_Money(rs(2))
            
            fLs.Cells(baris, 5).NumberFormat = "#,##0"
            
            fLs.Cells(baris, 6) = cek_Money(rs(3))
            fLs.Cells(baris, 6).NumberFormat = "#,##0"
            
            If cek_Money(rs(2)) > 0 Then
                fLs.Cells(baris, 7) = (cek_Money(rs(3)) / cek_Money(rs(2))) * 100
            Else
                fLs.Cells(baris, 7) = "-"
            End If
            
            fLs.Cells(baris, 8) = cek_Money(rs(5))
            fLs.Cells(baris, 8).NumberFormat = "#,##0"
            totalPBKonstruksi = totalPBKonstruksi + cek_Money(rs(5))
            
            fLs.Cells(baris, 9) = cek_Money(rs(6))
            fLs.Cells(baris, 9).NumberFormat = "#,##0"
            
            
            If cek_Money(rs(5)) > 0 Then
                fLs.Cells(baris, 10) = (cek_Money(rs(6)) / cek_Money(rs(5))) * 100
            Else
                fLs.Cells(baris, 10) = "-"
            End If
            
            fLs.Cells(baris, 11) = cek_Money(rs(8))
            fLs.Cells(baris, 11).NumberFormat = "#,##0"
            fLs.Cells(baris, 12) = cek_Money(rs(9))
            fLs.Cells(baris, 12).NumberFormat = "#,##0"
            fLs.Cells(baris, 13) = cek_null(rs(10))
            
            fLs.Cells(baris, 14) = cek_Money(rs(11))
            fLs.Cells(baris, 14).NumberFormat = "#,##0"
            fLs.Cells(baris, 15) = cek_Money(rs(12))
            fLs.Cells(baris, 15).NumberFormat = "#,##0"
            fLs.Cells(baris, 16) = cek_null(rs(13))
            
            fLs.Cells(baris, 17) = cek_Money(rs(14))
            fLs.Cells(baris, 17).NumberFormat = "#,##0"
            fLs.Cells(baris, 18) = cek_Money(rs(15))
            fLs.Cells(baris, 18).NumberFormat = "#,##0"
            
            If cek_Money(rs(14)) > 0 Then
                fLs.Cells(baris, 19) = (cek_Money(rs(15)) / cek_Money(rs(14))) * 100
            Else
                fLs.Cells(baris, 19) = "-"
            End If
            rs.MoveNext
            c = c + 1
            baris = baris + 1
        Loop
    End If
    
    'total
    fLs.Cells(baris, 5) = totalPBSewa
    fLs.Cells(baris, 5).NumberFormat = "#,##0"
    
    fLs.Cells(baris, 8) = totalPBKonstruksi
    fLs.Cells(baris, 8).NumberFormat = "#,##0"
    baris = baris + 2
    
    Call setListInfo(Me.List1, "Proses Sheet01 - get TB")
    
    'next part, LABEL nya
    data1 = Array("Data Akuntansi", "Biaya Upah(50101)", "Biaya Subkont(50301)", _
                "Biaya Subkont Beda Kurs ( 50302 )", _
                "Beban Sewa Gedung /Kantor ( 51851 )", _
                "Beban Sewa Gedung /Kantor ( 83251 )", _
                "", _
                "Hutang Awal :", "Hutang Awal Subkont ( 20111 )", _
                "Hutang Awal Subkont YBDF ( 20112 )", _
                "Hutang Awal Retensi Subkont ( 20113 )", _
                "Hutang Awal Subkont Beda Kurs ( 20116 )", _
                "Hutang Awal Upah ( 20131 )", _
                "", _
                "Hutang Akhir:", "Hutang Awal Subkont ( 20111 )", _
                "Hutang Awal Subkont YBDF ( 20112 )", _
                "Hutang Awal Retensi Subkont ( 20113 )", _
                "Hutang Awal Subkont Beda Kurs ( 20116 )", _
                "Hutang Awal Upah ( 20131 )", _
                "", _
                "Objek PPh 4(2) Konstruksi  Pada Akuntansi", _
                "Objek PPh 4(2) Sewa Pada Akuntansi", _
                "", _
                "Objek PPh 4(2) Konstruksi  Yg Belum di SPTkan", _
                "Objek PPh 4(2) Sewa  Yg Belum di SPTkan")
    'index kolom
    data2 = Array(0, 3, 3, 3, 0, 0, _
                0, _
                0, 3, 3, 3, 3, 3, _
                0, _
                0, 3, 3, 3, 3, 3, _
                0, _
                3, 0, _
                0, _
                3, 0)
                
    Total1 = get_nilaiAkunAll_divisi("50101", tahun, DIVISI) + _
               get_nilaiAkunAll_divisi("50301", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("50302", tahun, DIVISI) + _
                (get_nilaiAkunAll_divisi("20111", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20112", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20113", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20116", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20131", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20111", tahun, DIVISI) * 1) + _
                (get_nilaiAkunAll_divisi("20112", tahun, DIVISI) * 1) + _
                (get_nilaiAkunAll_divisi("20113", tahun, DIVISI) * 1) + _
                (get_nilaiAkunAll_divisi("20116", tahun, DIVISI) * 1) + _
                (get_nilaiAkunAll_divisi("20131", tahun, DIVISI) * 1)
                
    isi = Array("", get_nilaiAkunAll_divisi("50101", tahun, DIVISI), _
                get_nilaiAkunAll_divisi("50301", tahun, DIVISI), _
                get_nilaiAkunAll_divisi("50302", tahun, DIVISI), _
                get_nilaiAkunAll_divisi("51851", tahun, DIVISI), _
                get_nilaiAkunAll_divisi("83251", tahun, DIVISI), _
                "", "", _
                get_nilaiAkunAll_divisi("20111", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                get_nilaiAkunAll_divisi("20112", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                get_nilaiAkunAll_divisi("20113", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                get_nilaiAkunAll_divisi("20116", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                get_nilaiAkunAll_divisi("20131", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                "", "", _
                get_nilaiAkunAll_divisi("20111", tahun, DIVISI) * 1, _
                get_nilaiAkunAll_divisi("20112", tahun, DIVISI) * 1, _
                get_nilaiAkunAll_divisi("20113", tahun, DIVISI) * 1, _
                get_nilaiAkunAll_divisi("20116", tahun, DIVISI) * 1, _
                get_nilaiAkunAll_divisi("20131", tahun, DIVISI) * 1, _
                "", Total1, _
                get_nilaiAkunAll_divisi("51851", tahun, DIVISI) + get_nilaiAkunAll_divisi("83251", tahun, DIVISI), _
                "", _
                Total1 - totalPBKonstruksi, _
                (get_nilaiAkunAll_divisi("51851", tahun, DIVISI) + get_nilaiAkunAll_divisi("83251", tahun, DIVISI)) - totalPBSewa)
    
    'MsgBox UBound(data1)
    'MsgBox UBound(data2)
    'MsgBox UBound(isi)
    
    kolom = 3
    baris = baris + 2
    For c = 0 To UBound(data1)
        fLs.Cells(baris, kolom) = data1(c)
        fLs.Cells(baris, kolom + 2 + data2(c)) = isi(c)
        If Trim(isi(c)) = "" Then
        Else
            fLs.Cells(baris, kolom + 2 + data2(c)).NumberFormat = "#,##0"
        End If
        
        baris = baris + 1
    Next
    
    'fl.ActiveWorkbook.Save
    fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub

Sub pph21_sheet_01(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim nmFile As String, sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long
    
    Dim kryw_ttp_atas_ptkp_jml_Karyawan As Integer
    Dim kryw_ttp_atas_ptkp_Penghasilan_Bruto As Currency, kryw_ttp_atas_ptkp_PPh_21 As Currency
    Dim kryw_ttp_bwh_ptkp_jml_Karyawan As Integer, kryw_ttp_bwh_Penghasilan_Bruto As Currency
    Dim pesangon_jml_Karyawan As Integer, pesangon_Penghasilan_Bruto As Currency, pesangon_PPh_21 As Currency
    Dim total_Karyawan As Integer, total_Penghasilan_Bruto As Currency, total_PPh_21 As Currency
    Dim P1 As Double
    
    Dim SUM_total_Penghasilan_Bruto As Currency, SUM_total_PPh_21 As Currency
    Dim data1, isi, Total1 As Currency
    
    nmFile = App.Path & "\rep\uReport_KonsepRekapSPT21.xlsx"
    f = nmFile
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet01 - error open xls template")
            Exit Sub
        End If
    Else
        MsgBox "File template tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet01 - error open xls template")
        Exit Sub
    End If
    
    'open sheet 1, isi data looping
    Set fLs = fl.Sheets(1)
    baris = 2
    kolom = 2
    
    fLs.Cells(baris, kolom) = "REKAP SSP + SPM " & Me.cb_pph & " - Tahun " & tahun
    baris = 4
    fLs.Cells(baris, kolom) = "Filter Unit: " & DIVISI
    baris = 5
    fLs.Cells(baris, kolom) = "Filter Bulan: " & Masa_Pajak
    baris = 6
    fLs.Cells(baris, kolom) = "Filter KPP: " & kpp
    'fLs.Cells(baris, kolom + 3) = 2500000
    'fLs.Cells(baris, kolom + 3).NumberFormat = "#,##0"
                                    
    baris = 7
    kolom = 5
    fLs.Cells(baris, kolom) = "Data dari awal s/d Bulan:" & Masa_Pajak & " tahun " & _
                            tahun
    
    '-- get data per KPP
    sql = "Select mkpp.npwp, mkpp.kpp_administrasi, " & _
        "F_get_kryw_ttp_jml('2020','8',mkpp.npwp,'') as kryw_ttp_atas_ptkp_jml_Karyawan, " & _
        "F_get_kryw_ttp_bruto('2020','8',mkpp.npwp,'') as kryw_ttp_atas_ptkp_Penghasilan_Bruto, " & _
        "F_get_kryw_ttp_pph21('2020','8',mkpp.npwp,'') as kryw_ttp_atas_ptkp_PPh_21, " & _
        "F_get_kryw_ttp_bwh_ptkp_jml('2020','8',mkpp.npwp,'') as kryw_ttp_bwh_ptkp_jml_Karyawan, " & _
        "F_get_kryw_ttp_bwh_ptkp_bruto('2020','8',mkpp.npwp,'') as kryw_ttp_bwh_Penghasilan_Bruto, " & _
        "'0' as kryw_non_ttp_jml_Karyawan, " & _
        "'0' as kryw_non_ttp_Penghasilan_Bruto, " & _
        "'0' as kryw_non_ttp_PPh_21, " & _
        "F_get_pesangon_jml('2020','8',mkpp.npwp,'') as pesangon_jml_Karyawan, " & _
        "F_get_pesangon_bruto('2020','8',mkpp.npwp,'') as pesangon_Penghasilan_Bruto, " & _
        "F_get_pesangon_pph('2020','8',mkpp.npwp,'') as pesangon_PPh_21, " & _
        "'0' as total_Karyawan, " & _
        "'0' as total_Penghasilan_Bruto, " & _
        "'0' as total_PPh_21, " & _
        "'0' as P1 " & _
        "From mkpp  "
    If Trim(UCase(Me.cb_kpp.text)) = "ALL" Then
        sql = sql & "order by npwp"
    Else
        sql = sql & "where npwp = '" & get_kode_combo(Me.cb_kpp, "#") & "' " & _
        "order by npwp"
    End If
        
    
    Call setListInfo(Me.List1, "Proses Sheet01 - Open Query..")
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql errir", "", sql)
        Exit Sub
    End If
    
    Call setListInfo(Me.List1, "Proses Sheet01 - Open Query..OK")
    jRec = RecordCount(rs)
    
    '0,1: mkpp.npwp, mkpp.kpp_administrasi,
    '2: F_get_kryw_ttp_jml('2020','8',mkpp.npwp,'') as kryw_ttp_atas_ptkp_jml_Karyawan,
    '3: F_get_kryw_ttp_bruto('2020','8',mkpp.npwp,'') as kryw_ttp_atas_ptkp_Penghasilan_Bruto,
    '4: F_get_kryw_ttp_pph21('2020','8',mkpp.npwp,'') as kryw_ttp_atas_ptkp_PPh_21,

    '5: F_get_kryw_ttp_bwh_ptkp_jml('2020','8',mkpp.npwp,'') as kryw_ttp_bwh_ptkp_jml_Karyawan,
    '6: F_get_kryw_ttp_bwh_ptkp_bruto('2020','8',mkpp.npwp,'') as kryw_ttp_bwh_Penghasilan_Bruto,

    '7: '0' as kryw_non_ttp_jml_Karyawan,
    '8: '0' as kryw_non_ttp_Penghasilan_Bruto,
    '9: '0' as kryw_non_ttp_PPh_21,

    '10: F_get_pesangon_jml('2020','8',mkpp.npwp,'') as pesangon_jml_Karyawan,
    '11: F_get_pesangon_bruto('2020','8',mkpp.npwp,'') as pesangon_Penghasilan_Bruto,
    '12: F_get_pesangon_pph('2020','8',mkpp.npwp,'') as pesangon_PPh_21,

    '13: '0' as total_Karyawan,
    '14: '0' as total_Penghasilan_Bruto,
    '15: '0' as total_PPh_21,
    '16:  '0' as P1

    
    
    
    baris = 10
    If jRec > 0 Then
        rs.MoveFirst
        c = 1
        SUM_total_Penghasilan_Bruto = 0
        SUM_total_PPh_21 = 0
        Do While rs.EOF = False
            Call info_progress(Me.StatusBar1, 1, c, jRec, "write data perKPP")
            fLs.Cells(baris, 3) = "'" & cek_null(rs(0))
            fLs.Cells(baris, 4) = cek_null(rs(1))
            
            fLs.Cells(baris, 5) = cek_Int(rs(2))
            
            fLs.Cells(baris, 6) = cek_Money(rs(3))
            fLs.Cells(baris, 6).NumberFormat = "#,##0"
            
            fLs.Cells(baris, 7) = cek_Money(rs(4))
            fLs.Cells(baris, 7).NumberFormat = "#,##0"
                        
            fLs.Cells(baris, 8) = cek_Int(rs(5))
                        
            fLs.Cells(baris, 9) = cek_Money(rs(6))
            fLs.Cells(baris, 9).NumberFormat = "#,##0"
            
            fLs.Cells(baris, 10) = cek_Int(rs(7))
            
            fLs.Cells(baris, 11) = cek_Money(rs(8))
            fLs.Cells(baris, 11).NumberFormat = "#,##0"
            fLs.Cells(baris, 12) = cek_Money(rs(9))
            fLs.Cells(baris, 12).NumberFormat = "#,##0"
            
            fLs.Cells(baris, 13) = cek_Int(rs(10))
            
            fLs.Cells(baris, 14) = cek_Money(rs(11))
            fLs.Cells(baris, 14).NumberFormat = "#,##0"
            
            fLs.Cells(baris, 15) = cek_Money(rs(12))
            fLs.Cells(baris, 15).NumberFormat = "#,##0"
                        
            fLs.Cells(baris, 16) = cek_Int(rs(2)) + cek_Int(rs(5)) + cek_Int(rs(7)) + _
                                cek_Int(rs(10))
            
            total_Penghasilan_Bruto = cek_Money(rs(3)) + cek_Money(rs(6)) + cek_Money(rs(8)) + _
                                cek_Money(rs(11))
            fLs.Cells(baris, 17) = total_Penghasilan_Bruto
            fLs.Cells(baris, 17).NumberFormat = "#,##0"
            
            total_PPh_21 = cek_Money(rs(4)) + cek_Money(rs(7)) + cek_Money(rs(9)) + _
                                cek_Money(rs(12))
            fLs.Cells(baris, 18) = total_PPh_21
            fLs.Cells(baris, 18).NumberFormat = "#,##0"
            
            If total_Penghasilan_Bruto > 0 Then
                P1 = (total_PPh_21 / total_Penghasilan_Bruto) * 100
            Else
                P1 = 0
            End If
            fLs.Cells(baris, 19) = P1
            
            SUM_total_Penghasilan_Bruto = SUM_total_Penghasilan_Bruto + total_Penghasilan_Bruto
            SUM_total_PPh_21 = SUM_total_PPh_21 + total_PPh_21
            
            rs.MoveNext
            c = c + 1
            baris = baris + 1
        Loop
    End If
    
    'total
    fLs.Cells(baris, 17) = SUM_total_Penghasilan_Bruto
    fLs.Cells(baris, 17).NumberFormat = "#,##0"
    
    fLs.Cells(baris, 18) = SUM_total_PPh_21
    fLs.Cells(baris, 18).NumberFormat = "#,##0"
    
    fLs.Cells(baris + 1, 3) = ""
    fLs.Cells(baris + 1, 4) = ""
    fLs.Cells(baris + 2, 3) = ""
    fLs.Cells(baris + 2, 4) = ""
    fLs.Cells(baris + 3, 3) = ""
    fLs.Cells(baris + 3, 4) = ""
    fLs.Cells(baris + 4, 3) = ""
    fLs.Cells(baris + 4, 4) = ""
    
    baris = baris + 2
    
    Call setListInfo(Me.List1, "Proses Sheet01 - get TB")
    
    'next part, LABEL nya
    data1 = Array("Data Akuntansi", _
            "51101   Gaji", "51111   Tunjangan Fungsional", "51114   Tunjangan Penggantian Cuti", "51115   Thr", _
            "51116   Tunj Iuran Astek&Ass keclkaan", "51119   Tunj PPh Pasal 21", "51122   Biaya Pengobatan", "51125   Uang Pesangon", _
            "51201   Gaji", "51206   Tunjangan Transpor", "51213   Tunjangan Transpor", "51215   Thr", _
            "51216   Tunj Iuran Astek & Kecelakaan", "51219   Tunjangan PPh Pasal 21", "51221   Uang Lembur dan Transpor", "51222   Biaya Pengobatan", _
            "51225   Uang Pesangon", "51228   Insentif PBK", "51501   Beban Perjalanan Dinas", "51502   Beban Pendidikan", _
            "51861   Bbn Pakaian Dns & Perlk Kerja", "80101   Gaji PEGAWAI", "80111   Tunjangan Fungsional Pegawai", "80114   Tunjangan Penggantian Cuti Peg", _
            "80115   Thr", "80116   Tunj Iuran ASTEK & Ass Keclkn", "80118   Tunjangan Variable", "80119   Tunjangan PPh Pasal 21  Pegawa", _
            "80121   Uang Lembur dan Transpor Peg", "80123   Beban Pengobatan", "80124   PREMI SATYA JASA", "80125   Uang Pesangon", _
            "80126   Gaji PBT Idle", "80127   Kenaikan Gaji Pegawai", "80128   Insentif PEGAWAI", "80131   GAJI  DIREKSI DEKOM", _
            "80133   TUNJ.PPH DIR DEKOM", "80136   THR DIR KOM", "80137   Tantiem 2007", "80138   Gaji Komisaris", _
            "80139   Tunjangan PPh Komisaris", "80201   Gaji", "80214   Tunjangan Penggantian Cuti", "80215   Tunjangan Hari Raya (THR)", _
            "80216   Tunj Iuran Astek & Ass Kecl", "80219   Tunjangan PPh Pasal 21", "80221   Uang Lembur dan Transpor", "80223   Biaya Pengobatan", _
            "80225   Uang Pesangon", "80501   Beban Pendidikan", "80601   Beban Perjalanan Dinas", "83261   Beban Pakaian Seragam", _
            "", "Hutang Awal:", "20701   Hutang Gaji", "20704   Hutang Insentif", "20709   Hutang Insentif", _
            "21902   Hutang Tantiem", "", "Hutang Akhir:", "20701   Hutang Gaji", "20704   Hutang Insentif", _
            "20709   Hutang Insentif", "21902   Hutang Tantiem", "", "Objek PPh 21 Pada Akuntansi", _
            "", "Objek PPh 21 Yang Belum diSPTkan", "( B-A )")
                
    Total1 = get_nilaiAkunAll_divisi("51101", tahun, DIVISI) + get_nilaiAkunAll_divisi("51111", tahun, DIVISI) + get_nilaiAkunAll_divisi("51114", tahun, DIVISI) + get_nilaiAkunAll_divisi("51115", tahun, DIVISI) + get_nilaiAkunAll_divisi("51116", tahun, DIVISI) + get_nilaiAkunAll_divisi("51119", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("51122", tahun, DIVISI) + get_nilaiAkunAll_divisi("51125", tahun, DIVISI) + get_nilaiAkunAll_divisi("51201", tahun, DIVISI) + get_nilaiAkunAll_divisi("51206", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("51213", tahun, DIVISI) + get_nilaiAkunAll_divisi("51215", tahun, DIVISI) + get_nilaiAkunAll_divisi("51216", tahun, DIVISI) + get_nilaiAkunAll_divisi("51219", tahun, DIVISI) + get_nilaiAkunAll_divisi("51221", tahun, DIVISI) + get_nilaiAkunAll_divisi("51222", tahun, DIVISI) + get_nilaiAkunAll_divisi("51225", tahun, DIVISI) + get_nilaiAkunAll_divisi("51228", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("51501", tahun, DIVISI) + get_nilaiAkunAll_divisi("51502", tahun, DIVISI) + get_nilaiAkunAll_divisi("51861", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("80101", tahun, DIVISI) + get_nilaiAkunAll_divisi("80111", tahun, DIVISI) + get_nilaiAkunAll_divisi("80114", tahun, DIVISI) + get_nilaiAkunAll_divisi("80115", tahun, DIVISI) + get_nilaiAkunAll_divisi("80116", tahun, DIVISI) + get_nilaiAkunAll_divisi("80118", tahun, DIVISI) + get_nilaiAkunAll_divisi("80119", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("80121", tahun, DIVISI) + get_nilaiAkunAll_divisi("80123", tahun, DIVISI) + get_nilaiAkunAll_divisi("80124", tahun, DIVISI) + get_nilaiAkunAll_divisi("80125", tahun, DIVISI) + get_nilaiAkunAll_divisi("80126", tahun, DIVISI) + get_nilaiAkunAll_divisi("80127", tahun, DIVISI) + get_nilaiAkunAll_divisi("80128", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("80131", tahun, DIVISI) + get_nilaiAkunAll_divisi("80133", tahun, DIVISI) + get_nilaiAkunAll_divisi("80136", tahun, DIVISI) + get_nilaiAkunAll_divisi("80137", tahun, DIVISI) + get_nilaiAkunAll_divisi("80138", tahun, DIVISI) + get_nilaiAkunAll_divisi("80139", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("80201", tahun, DIVISI) + get_nilaiAkunAll_divisi("80214", tahun, DIVISI) + get_nilaiAkunAll_divisi("80215", tahun, DIVISI) + get_nilaiAkunAll_divisi("80216", tahun, DIVISI) + get_nilaiAkunAll_divisi("80219", tahun, DIVISI) + get_nilaiAkunAll_divisi("80221", tahun, DIVISI) + get_nilaiAkunAll_divisi("80223", tahun, DIVISI) + get_nilaiAkunAll_divisi("80225", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("80501", tahun, DIVISI) + get_nilaiAkunAll_divisi("80601", tahun, DIVISI) + get_nilaiAkunAll_divisi("83261", tahun, DIVISI) + _
                (get_nilaiAkunAll_divisi("20701", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20704", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20709", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("21902", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20701", tahun, DIVISI) * 1) + _
                (get_nilaiAkunAll_divisi("20704", tahun, DIVISI) * 1) + _
                (get_nilaiAkunAll_divisi("20709", tahun, DIVISI) * 1) + _
                (get_nilaiAkunAll_divisi("21902", tahun, DIVISI) * 1)
                
                
                
    isi = Array("", get_nilaiAkunAll_divisi("51101", tahun, DIVISI), get_nilaiAkunAll_divisi("51111", tahun, DIVISI), get_nilaiAkunAll_divisi("51114", tahun, DIVISI), get_nilaiAkunAll_divisi("51115", tahun, DIVISI), get_nilaiAkunAll_divisi("51116", tahun, DIVISI), get_nilaiAkunAll_divisi("51119", tahun, DIVISI), _
                get_nilaiAkunAll_divisi("51122", tahun, DIVISI), get_nilaiAkunAll_divisi("51125", tahun, DIVISI), get_nilaiAkunAll_divisi("51201", tahun, DIVISI), get_nilaiAkunAll_divisi("51206", tahun, DIVISI), _
                get_nilaiAkunAll_divisi("51213", tahun, DIVISI), get_nilaiAkunAll_divisi("51215", tahun, DIVISI), get_nilaiAkunAll_divisi("51216", tahun, DIVISI), get_nilaiAkunAll_divisi("51219", tahun, DIVISI), get_nilaiAkunAll_divisi("51221", tahun, DIVISI), get_nilaiAkunAll_divisi("51222", tahun, DIVISI), get_nilaiAkunAll_divisi("51225", tahun, DIVISI), get_nilaiAkunAll_divisi("51228", tahun, DIVISI), _
                get_nilaiAkunAll_divisi("51501", tahun, DIVISI), get_nilaiAkunAll_divisi("51502", tahun, DIVISI), get_nilaiAkunAll_divisi("51861", tahun, DIVISI), _
                get_nilaiAkunAll_divisi("80101", tahun, DIVISI), get_nilaiAkunAll_divisi("80111", tahun, DIVISI), get_nilaiAkunAll_divisi("80114", tahun, DIVISI), get_nilaiAkunAll_divisi("80115", tahun, DIVISI), get_nilaiAkunAll_divisi("80116", tahun, DIVISI), get_nilaiAkunAll_divisi("80118", tahun, DIVISI), get_nilaiAkunAll_divisi("80119", tahun, DIVISI), _
                get_nilaiAkunAll_divisi("80121", tahun, DIVISI), get_nilaiAkunAll_divisi("80123", tahun, DIVISI), get_nilaiAkunAll_divisi("80124", tahun, DIVISI), get_nilaiAkunAll_divisi("80125", tahun, DIVISI), get_nilaiAkunAll_divisi("80126", tahun, DIVISI), get_nilaiAkunAll_divisi("80127", tahun, DIVISI), get_nilaiAkunAll_divisi("80128", tahun, DIVISI), _
                get_nilaiAkunAll_divisi("80131", tahun, DIVISI), get_nilaiAkunAll_divisi("80133", tahun, DIVISI), get_nilaiAkunAll_divisi("80136", tahun, DIVISI), get_nilaiAkunAll_divisi("80137", tahun, DIVISI), get_nilaiAkunAll_divisi("80138", tahun, DIVISI), get_nilaiAkunAll_divisi("80139", tahun, DIVISI), _
                get_nilaiAkunAll_divisi("80201", tahun, DIVISI), get_nilaiAkunAll_divisi("80214", tahun, DIVISI), get_nilaiAkunAll_divisi("80215", tahun, DIVISI), get_nilaiAkunAll_divisi("80216", tahun, DIVISI), get_nilaiAkunAll_divisi("80219", tahun, DIVISI), get_nilaiAkunAll_divisi("80221", tahun, DIVISI), get_nilaiAkunAll_divisi("80223", tahun, DIVISI), get_nilaiAkunAll_divisi("80225", tahun, DIVISI), _
                get_nilaiAkunAll_divisi("80501", tahun, DIVISI), get_nilaiAkunAll_divisi("80601", tahun, DIVISI), get_nilaiAkunAll_divisi("83261", tahun, DIVISI), _
                "", "", _
                get_nilaiAkunAll_divisi("20701", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                get_nilaiAkunAll_divisi("20704", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                get_nilaiAkunAll_divisi("20709", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                get_nilaiAkunAll_divisi("21902", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                "", "", _
                get_nilaiAkunAll_divisi("20701", tahun, DIVISI) * 1, get_nilaiAkunAll_divisi("20704", tahun, DIVISI) * 1, get_nilaiAkunAll_divisi("20709", tahun, DIVISI) * 1, get_nilaiAkunAll_divisi("21902", tahun, DIVISI) * 1, "", "", Total1, _
                "", Total1 - SUM_total_Penghasilan_Bruto)
    
    'MsgBox UBound(data1)
    'MsgBox UBound(data2)
    'MsgBox UBound(isi)
    
    kolom = 3
    fLs.Cells(baris + 1, 1) = ""
    fLs.Cells(baris + 1, 2) = ""
    fLs.Cells(baris + 2, 1) = ""
    fLs.Cells(baris + 2, 2) = ""
    baris = baris + 2
    For c = 0 To UBound(data1)
        fLs.Cells(baris, kolom) = data1(c)
        fLs.Cells(baris, kolom + 1) = ""
        fLs.Cells(baris, kolom + 2) = isi(c)
        If Trim(isi(c)) = "" Then
        Else
            fLs.Cells(baris, kolom + 2).NumberFormat = "#,##0"
        End If
        baris = baris + 1
    Next
    
    'fl.ActiveWorkbook.Save
    fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub


Sub pph22_sheet_01(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim nmFile As String, sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long
    
    Dim SUM_total_Penghasilan_Bruto As Currency, SUM_total_PPh_22 As Currency
    Dim data1, isi, Total1 As Currency
    
    nmFile = App.Path & "\rep\uReport_KonsepRekapSPT22.xls"
                              
    f = nmFile
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet01 - error open xls template")
            Exit Sub
        End If
    Else
        MsgBox "File template tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet01 - error open xls template")
        Exit Sub
    End If
    
    'open sheet 1, isi data looping
    Set fLs = fl.Sheets(1)
    baris = 2
    kolom = 2
    
    fLs.Cells(baris, kolom) = "REKAP SSP + SPM " & Me.cb_pph & " - Tahun " & tahun
    baris = 4
    fLs.Cells(baris, kolom) = "Filter Unit: " & DIVISI
    baris = 5
    fLs.Cells(baris, kolom) = "Filter Bulan: " & Masa_Pajak
    baris = 6
    fLs.Cells(baris, kolom) = "Filter KPP: " & kpp
    'fLs.Cells(baris, kolom + 3) = 2500000
    'fLs.Cells(baris, kolom + 3).NumberFormat = "#,##0"
                                    
    baris = 8
    kolom = 5
    fLs.Cells(baris, kolom) = "Data dari awal s/d Bulan:" & Masa_Pajak & " tahun " & _
                            tahun
    
    '-- get data per KPP
    sql = "Select mkpp.npwp, mkpp.kpp_administrasi," & _
            "'PO' as SPT_Terakhir, '0' as jml_Industri_Eksportir, " & _
            "F_get_pph22_bruto('2020','12',mkpp.npwp,'') as jml_BUMN, " & _
            "'' as jml_Nilai_Objek_PPh_22, '' as jml_PPh_Industri_Eksportir, " & _
            "F_get_pph22_pph('2020','12',mkpp.npwp,'') as jml_PPh_22_BUMN, " & _
            "'' as jml_Total_PPh_22, '' as p1 " & _
            "From mkpp "
    If Trim(UCase(Me.cb_kpp.text)) = "ALL" Then
        sql = sql & "order by npwp"
    Else
        sql = sql & "where npwp = '" & get_kode_combo(Me.cb_kpp, "#") & "' " & _
        "order by npwp"
    End If
        
    
    Call setListInfo(Me.List1, "Proses Sheet01 - Open Query..")
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql errir", "", sql)
        Exit Sub
    End If
    
    Call setListInfo(Me.List1, "Proses Sheet01 - Open Query..OK")
    jRec = RecordCount(rs)
    
    '0: mkpp.npwp, mkpp.kpp_administrasi, 'PO' as SPT_Terakhir,
    '3: '0' as jml_Industri_Eksportir, F_get_pph22_bruto('2020','12',mkpp.npwp,'') as jml_BUMN,
    '5: '' as jml_Nilai_Objek_PPh_22, '' as jml_PPh_Industri_Eksportir,
    '7: F_get_pph22_pph('2020','12',mkpp.npwp,'') as jml_PPh_22_BUMN,
    '8: '' as jml_Total_PPh_22, '' as p1

    baris = 10
    If jRec > 0 Then
        rs.MoveFirst
        c = 1
        SUM_total_Penghasilan_Bruto = 0
        SUM_total_PPh_22 = 0
        Do While rs.EOF = False
            Call info_progress(Me.StatusBar1, 1, c, jRec, "write data perKPP")
            fLs.Cells(baris, 3) = "'" & cek_null(rs(0))
            fLs.Cells(baris, 4) = cek_null(rs(1))
            
            fLs.Cells(baris, 5) = cek_Int(rs(2))
            
            fLs.Cells(baris, 6) = cek_Money(rs(3))
            fLs.Cells(baris, 6).NumberFormat = "#,##0"
            
            fLs.Cells(baris, 7) = cek_Money(rs(4))
            fLs.Cells(baris, 7).NumberFormat = "#,##0"
                        
            fLs.Cells(baris, 8) = cek_Money(rs(3)) + cek_Money(rs(4))
            fLs.Cells(baris, 8).NumberFormat = "#,##0"
            SUM_total_Penghasilan_Bruto = SUM_total_Penghasilan_Bruto + _
                                        cek_Money(rs(3)) + cek_Money(rs(4))
                        
            fLs.Cells(baris, 9) = cek_Money(rs(6))
            fLs.Cells(baris, 9).NumberFormat = "#,##0"
            
            fLs.Cells(baris, 10) = cek_Money(rs(7))
            fLs.Cells(baris, 10).NumberFormat = "#,##0"
            
            fLs.Cells(baris, 11) = cek_Money(rs(6)) + cek_Money(rs(7))
            fLs.Cells(baris, 11).NumberFormat = "#,##0"
            SUM_total_PPh_22 = SUM_total_PPh_22 + cek_Money(rs(6)) + cek_Money(rs(7))
            
            If (cek_Money(rs(3)) + cek_Money(rs(4))) > 0 Then
                fLs.Cells(baris, 12) = (cek_Money(rs(6)) + cek_Money(rs(7))) / _
                                        (cek_Money(rs(3)) + cek_Money(rs(4)))
            Else
                fLs.Cells(baris, 12) = 0
            End If
            
            rs.MoveNext
            c = c + 1
            baris = baris + 1
        Loop
    End If
    
    'total
    fLs.Cells(baris, 8) = SUM_total_Penghasilan_Bruto
    fLs.Cells(baris, 8).NumberFormat = "#,##0"
    
    fLs.Cells(baris, 11) = SUM_total_PPh_22
    fLs.Cells(baris, 11).NumberFormat = "#,##0"
    
    fLs.Cells(baris + 1, 3) = ""
    fLs.Cells(baris + 1, 4) = ""
    fLs.Cells(baris + 2, 3) = ""
    fLs.Cells(baris + 2, 4) = ""
    fLs.Cells(baris + 3, 3) = ""
    fLs.Cells(baris + 3, 4) = ""
    fLs.Cells(baris + 4, 3) = ""
    fLs.Cells(baris + 4, 4) = ""
    
    baris = baris + 2
    
    Call setListInfo(Me.List1, "Proses Sheet01 - get TB")
    
    'next part, LABEL nya
    data1 = Array("Data Akuntansi", _
            "Biaya Bahan ( 50201 )", "Biaya Bahan Beda Kurs ( 50202 )", _
            "", "Hutang Awal:", "Hutang Supplier ( 20101 )", "Hutang Supplier YBDF ( 20102 )", _
            "", "Hutang Akhir:", "Hutang Supplier ( 20101 )", "Hutang Supplier YBDF ( 20102 )", _
            "", "Objek PPh 22 Pada Akuntansi", _
            "", "Objek PPh 22 Yang Belum diSPTkan", "( B-A )")
                
    Total1 = get_nilaiAkunAll_divisi("50201", tahun, DIVISI) + get_nilaiAkunAll_divisi("50202", tahun, DIVISI) + _
                (get_nilaiAkunAll_divisi("20101", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20102", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20101", tahun, DIVISI) * 1) + _
                (get_nilaiAkunAll_divisi("20102", tahun, DIVISI) * 1)
                
    isi = Array("", get_nilaiAkunAll_divisi("50201", tahun, DIVISI), get_nilaiAkunAll_divisi("50202", tahun, DIVISI), _
                "", "", _
                get_nilaiAkunAll_divisi("20101", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                get_nilaiAkunAll_divisi("20102", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                "", "", _
                get_nilaiAkunAll_divisi("20101", tahun, DIVISI) * 1, get_nilaiAkunAll_divisi("20102", tahun, DIVISI) * 1, _
                "", Total1, "", Total1 - SUM_total_Penghasilan_Bruto, "")
    
    'MsgBox UBound(data1)
    'MsgBox UBound(data2)
    'MsgBox UBound(isi)
    
    kolom = 3
    fLs.Cells(baris + 1, 1) = ""
    fLs.Cells(baris + 1, 2) = ""
    fLs.Cells(baris + 2, 1) = ""
    fLs.Cells(baris + 2, 2) = ""
    baris = baris + 2
    For c = 0 To UBound(data1)
        fLs.Cells(baris, kolom) = data1(c)
        fLs.Cells(baris, kolom + 1) = ""
        fLs.Cells(baris, kolom + 2) = isi(c)
        If Trim(isi(c)) = "" Then
        Else
            fLs.Cells(baris, kolom + 2).NumberFormat = "#,##0"
        End If
        baris = baris + 1
    Next
    
    'fl.ActiveWorkbook.Save
    fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub

Sub pph23_sheet_01(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim nmFile As String, sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long
    
    Dim SUM_total_Bruto As Currency, SUM_total_PPh_23 As Currency
    Dim data1, isi, Total1 As Currency
    
    nmFile = App.Path & "\rep\uReport_KonsepRekapSPT23.xls"
                              
    f = nmFile
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet01 - error open xls template")
            Exit Sub
        End If
    Else
        MsgBox "File template tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet01 - error open xls template")
        Exit Sub
    End If
    
    'open sheet 1, isi data looping
    Set fLs = fl.Sheets(1)
    baris = 2
    kolom = 2
    
    fLs.Cells(baris, kolom) = "REKAP SSP + SPM " & Me.cb_pph & " - Tahun " & tahun
    baris = 4
    fLs.Cells(baris, kolom) = "Filter Unit: " & DIVISI
    baris = 5
    fLs.Cells(baris, kolom) = "Filter Bulan: " & Masa_Pajak
    baris = 6
    fLs.Cells(baris, kolom) = "Filter KPP: " & kpp
    'fLs.Cells(baris, kolom + 3) = 2500000
    'fLs.Cells(baris, kolom + 3).NumberFormat = "#,##0"
                                    
    baris = 8
    kolom = 5
    fLs.Cells(baris, kolom) = "Data dari awal s/d Bulan:" & Masa_Pajak & " tahun " & _
                            tahun
    
    '-- get data per KPP
    sql = "Select mkpp.npwp, mkpp.kpp_administrasi, " & _
        "F_get_pendptan_sewa_bruto('2020','12',mkpp.npwp,'') as PB_Sewa, " & _
        "F_get_ssp_pph23_sewa('2020','12',mkpp.npwp,'') as SSP_Sewa, " & _
        "F_get_pendptan_jasa_bruto('2020','12',mkpp.npwp,'') as PB_Jasa, " & _
        "F_get_ssp_pph23_jasa('2020','12',mkpp.npwp,'') as SSP_Jasa, " & _
        "F_get_pendptan_LN_bruto('2020','12',mkpp.npwp,'') as PB_Luar_Negeri, " & _
        "F_get_ssp_pph('2020','12',mkpp.npwp,'','26') as SSP_PPh_26, " & _
        "'0' as Total_PB, " & _
        "'0' as Total_SSP, '' as p1 " & _
        "From mkpp "
    If Trim(UCase(Me.cb_kpp.text)) = "ALL" Then
        sql = sql & "order by npwp"
    Else
        sql = sql & "where npwp = '" & get_kode_combo(Me.cb_kpp, "#") & "' " & _
        "order by npwp"
    End If
        
    
    Call setListInfo(Me.List1, "Proses Sheet01 - Open Query..")
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql errir", "", sql)
        Exit Sub
    End If
    
    Call setListInfo(Me.List1, "Proses Sheet01 - Open Query..OK")
    jRec = RecordCount(rs)
    
    '0: mkpp.npwp, mkpp.kpp_administrasi,
    '2: F_get_pendptan_sewa_bruto('2020','12',mkpp.npwp,'') as PB_Sewa,
    '3: F_get_ssp_pph23_sewa('2020','12',mkpp.npwp,'') as SSP_Sewa,
    '4: F_get_pendptan_jasa_bruto('2020','12',mkpp.npwp,'') as PB_Jasa,
    '5: F_get_ssp_pph23_jasa('2020','12',mkpp.npwp,'') as SSP_Jasa,
    '6: F_get_pendptan_LN_bruto('2020','12',mkpp.npwp,'') as PB_Luar_Negeri,
    '7: F_get_ssp_pph('2020','12',mkpp.npwp,'','26') as SSP_PPh_26,
    '8: '0' as Total_PB, '0' as Total_SSP, '' as p1


    baris = 10
    If jRec > 0 Then
        rs.MoveFirst
        c = 1
        SUM_total_Bruto = 0
        SUM_total_PPh_23 = 0
        Do While rs.EOF = False
            Call info_progress(Me.StatusBar1, 1, c, jRec, "write data perKPP")
            fLs.Cells(baris, 3) = "'" & cek_null(rs(0))
            fLs.Cells(baris, 4) = cek_null(rs(1))
            
            fLs.Cells(baris, 5) = cek_Money(rs(2))
            fLs.Cells(baris, 5).NumberFormat = "#,##0"
            
            fLs.Cells(baris, 6) = cek_Money(rs(3))
            fLs.Cells(baris, 6).NumberFormat = "#,##0"
            
            fLs.Cells(baris, 7) = cek_Money(rs(4))
            fLs.Cells(baris, 7).NumberFormat = "#,##0"
                        
            fLs.Cells(baris, 8) = cek_Money(rs(5))
            fLs.Cells(baris, 8).NumberFormat = "#,##0"
            
            fLs.Cells(baris, 9) = cek_Money(rs(6))
            fLs.Cells(baris, 9).NumberFormat = "#,##0"
            
            fLs.Cells(baris, 10) = cek_Money(rs(7))
            fLs.Cells(baris, 10).NumberFormat = "#,##0"
            
            fLs.Cells(baris, 11) = cek_Money(rs(2)) + cek_Money(rs(4)) + cek_Money(rs(6))
            fLs.Cells(baris, 11).NumberFormat = "#,##0"
            
            fLs.Cells(baris, 12) = cek_Money(rs(3)) + cek_Money(rs(5)) + cek_Money(rs(7))
            fLs.Cells(baris, 12).NumberFormat = "#,##0"
            
            SUM_total_Bruto = SUM_total_Bruto + _
                                        cek_Money(rs(2)) + cek_Money(rs(4)) + cek_Money(rs(6))
            
            SUM_total_PPh_23 = SUM_total_PPh_23 + cek_Money(rs(3)) + cek_Money(rs(5)) + cek_Money(rs(7))
            
            If (cek_Money(rs(2)) + cek_Money(rs(4)) + cek_Money(rs(6))) > 0 Then
                fLs.Cells(baris, 13) = (cek_Money(rs(3)) + cek_Money(rs(5)) + cek_Money(rs(7))) / _
                                        (cek_Money(rs(2)) + cek_Money(rs(4)) + cek_Money(rs(6)))
            Else
                fLs.Cells(baris, 13) = 0
            End If
            
            rs.MoveNext
            c = c + 1
            baris = baris + 1
        Loop
    End If
    
    'total
    fLs.Cells(baris, 8) = SUM_total_Bruto
    fLs.Cells(baris, 8).NumberFormat = "#,##0"
    
    fLs.Cells(baris, 11) = SUM_total_PPh_23
    fLs.Cells(baris, 11).NumberFormat = "#,##0"
    
    fLs.Cells(baris + 1, 3) = ""
    fLs.Cells(baris + 1, 4) = ""
    fLs.Cells(baris + 2, 3) = ""
    fLs.Cells(baris + 2, 4) = ""
    fLs.Cells(baris + 3, 3) = ""
    fLs.Cells(baris + 3, 4) = ""
    fLs.Cells(baris + 4, 3) = ""
    fLs.Cells(baris + 4, 4) = ""
    
    baris = baris + 2
    
    Call setListInfo(Me.List1, "Proses Sheet01 - get TB")
    
    'next part, LABEL nya
    data1 = Array("Data Akuntansi", "50103 - Beban Upah Transporter", "50401 - Beban Sewa Alat Berat", "50402 - Beban Sewa Alat Ringan", "50405 - Bbn Sewa Alt Brt Dr Cab Prlt", "50406 - MR AB IDLE", "50411 - Bbn Pemeliharaan Alat Berat", "50412 - Bbn Pemeliharaan Alat Ringan", "50413 - Beban Sparepart Alat Berat", "50414 - Beban Sparepart Alat Ringan", "50431 - Beban Mob dan Demob Alat", "50432 - Biaya Operasional Alat Berat", "51401 - Beban Pemeliharaan Kendaraan", "51402 - Beban Pemel Inventaris Kantor", "51404 - Beban Pemel Bangunan", "51801 - Beban Kons.Jasa Manjn/Akt", "51802 - Beban Konsultasi Jasa Hukum", "51803 - Beban Pengawas Lapangan", "51852 - Beban Sewa Kendaraan", "51853 - Beban Sewa Komputer", "51901 - Reklame dan Advertensi", "51903 - Maket, Eksibisi & Sign Board", "71101 - Reklame dan Advertensi", "71102 - BROSUR DAN LEAFLET", "71103 - Maket, Eksibisi & Sign Board", "81104 - Bbn Pemel & Perb Bang & Pras", "81121 - BEBAN PEMEL & PERB KENDARAAN", _
            "81122 - Beban Pemel & Perb Inv.Kantor", "83107 - Beban Internet", "83211 - Beban Konsultasi Jasa Manajeme", "83212 - Beban Konsultasi Jasa Hukum", "83252 - Beban Sewa Kendaraan", "83253 - Beban Sewa Komputer", _
            "", "Hutang Awal:", "20133 - Hutang Pihak Ke 3 Pemeliharaan dan Perbaikan", "20138 - Hutang Transporter", _
            "", "Hutang Akhir:", "20133 - Hutang Pihak Ke 3 Pemeliharaan dan Perbaikan", "20138 - Hutang Transporter", _
            "", "Objek PPh 23 Pada Akuntansi", _
            "", "Objek PPh 23 Yang Belum diSPTkan", "( B-A )")
                
    Total1 = get_nilaiAkunAll_divisi("50103", tahun, DIVISI) + get_nilaiAkunAll_divisi("50401", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("50402", tahun, DIVISI) + get_nilaiAkunAll_divisi("50405", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("50406", tahun, DIVISI) + get_nilaiAkunAll_divisi("50411", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("50412", tahun, DIVISI) + get_nilaiAkunAll_divisi("50413", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("50414", tahun, DIVISI) + get_nilaiAkunAll_divisi("50431", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("50432", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("51401", tahun, DIVISI) + get_nilaiAkunAll_divisi("51402", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("51404", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("51801", tahun, DIVISI) + get_nilaiAkunAll_divisi("51802", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("51803", tahun, DIVISI) + get_nilaiAkunAll_divisi("51852", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("51853", tahun, DIVISI) + get_nilaiAkunAll_divisi("51901", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("51903", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("71101", tahun, DIVISI) + get_nilaiAkunAll_divisi("71102", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("71103", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("81104", tahun, DIVISI) + get_nilaiAkunAll_divisi("81121", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("81122", tahun, DIVISI) + get_nilaiAkunAll_divisi("83107", tahun, DIVISI) + get_nilaiAkunAll_divisi("83211", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("83212", tahun, DIVISI) + get_nilaiAkunAll_divisi("83252", tahun, DIVISI) + get_nilaiAkunAll_divisi("83253", tahun, DIVISI) + _
                (get_nilaiAkunAll_divisi("20133", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20138", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20133", tahun, DIVISI) * 1) + _
                (get_nilaiAkunAll_divisi("20138", tahun, DIVISI) * 1)
                
    isi = Array("", get_nilaiAkunAll_divisi("50103", tahun, DIVISI), get_nilaiAkunAll_divisi("50401", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("50402", tahun, DIVISI), get_nilaiAkunAll_divisi("50405", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("50406", tahun, DIVISI), get_nilaiAkunAll_divisi("50411", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("50412", tahun, DIVISI), get_nilaiAkunAll_divisi("50413", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("50414", tahun, DIVISI), get_nilaiAkunAll_divisi("50431", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("50432", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("51401", tahun, DIVISI), get_nilaiAkunAll_divisi("51402", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("51404", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("51801", tahun, DIVISI), get_nilaiAkunAll_divisi("51802", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("51803", tahun, DIVISI), get_nilaiAkunAll_divisi("51852", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("51853", tahun, DIVISI), get_nilaiAkunAll_divisi("51901", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("51903", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("71101", tahun, DIVISI), get_nilaiAkunAll_divisi("71102", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("71103", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("81104", tahun, DIVISI), get_nilaiAkunAll_divisi("81121", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("81122", tahun, DIVISI), get_nilaiAkunAll_divisi("83107", tahun, DIVISI), get_nilaiAkunAll_divisi("83211", tahun, DIVISI), _
         get_nilaiAkunAll_divisi("83212", tahun, DIVISI), get_nilaiAkunAll_divisi("83252", tahun, DIVISI), get_nilaiAkunAll_divisi("83253", tahun, DIVISI), _
                "", "", _
                get_nilaiAkunAll_divisi("20133", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                get_nilaiAkunAll_divisi("20138", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                "", "", _
                get_nilaiAkunAll_divisi("20133", tahun, DIVISI) * 1, get_nilaiAkunAll_divisi("20138", tahun, DIVISI) * 1, _
                "", Total1, "", Total1 - SUM_total_Bruto, "")
    
    'MsgBox UBound(data1)
    'MsgBox UBound(data2)
    'MsgBox UBound(isi)
    
    kolom = 3
    fLs.Cells(baris + 1, 1) = ""
    fLs.Cells(baris + 1, 2) = ""
    fLs.Cells(baris + 2, 1) = ""
    fLs.Cells(baris + 2, 2) = ""
    baris = baris + 2
    For c = 0 To UBound(data1)
        fLs.Cells(baris, kolom) = data1(c)
        fLs.Cells(baris, kolom + 1) = ""
        fLs.Cells(baris, kolom + 2) = isi(c)
        If Trim(isi(c)) = "" Then
        Else
            fLs.Cells(baris, kolom + 2).NumberFormat = "#,##0"
        End If
        baris = baris + 1
    Next
    
    'fl.ActiveWorkbook.Save
    fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub

Function get_pelaporan_spt_masa(DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String) As Recordset
    
    Dim sql As String
    Dim rs As Recordset
    
    If Trim(UCase(DIVISI)) = "ALL" Then
        DIVISI = ""
    End If
    
    '-- looping mengisi KPP
    sql = "Select mkpp.npwp, mkpp.kpp_administrasi, " & _
        "'P0' as SPT_Terakhir, " & _
        "F_get_pph42_sewa_bruto_bln('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp,'" & DIVISI & "') as s_Sewa_Tanah_Bangunan_bruto, " & _
        "F_get_pph42_konstruksi_bruto_bln('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp,'" & DIVISI & "') as s_Jasa_Konstruksi_bruto, " & _
        "F_get_pph42_obligasi_bruto_bln('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp,'" & DIVISI & "') as s_Bunga_Obligasi_bruto, " & _
        "'' as s_Deviden, " & _
        "'jumlah' as s_Nilai_Objek_PPh_42, " & _
        "F_get_ssp_pph('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp, '" & DIVISI & "','PPH FINAL SEWA') as s_PPh_Sewa_Tanah_Bangunan_ssp, " & _
        "'' as p1, " & _
        "F_get_ssp_pph('" & tahun & "','" & Masa_Pajak & "',mkpp.npwp, '" & DIVISI & "','PPH FINAL') as s_PPh_Jasa_Konstruksi_ssp, " & _
        "'' as p2, " & _
        "'' as s_PPh_Bunga_Obligasi_ssp, " & _
        "'' as p3, " & _
        "'' as s_PPh_Deviden_ssp, " & _
        "'' as p4, " & _
        "'' as s_Total_PPh_42_ssp " & _
        "From mkpp order by npwp "
    'sql = InputBox("sql error", "", sql)
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Set get_pelaporan_spt_masa = Null
    Else
        Set get_pelaporan_spt_masa = rs
    End If
End Function

Sub pph42_sheet_02(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long, bln As Integer
    Dim Total1 As Currency, totalPBKonstruksi As Currency, totalPBSewa As Currency
    Dim data1, data2, isi
    
    f = fileSimpan
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet02 - file not found")
            Exit Sub
        End If
    Else
        MsgBox "File tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet02 - file not found")
            Exit Sub
        Exit Sub
    End If
    
    'open sheet 2, isi data looping
    Set fLs = fl.Sheets(2)
    baris = 2
    kolom = 2
    
    fLs.Cells(baris, kolom) = "REKAP PELAPORAN SPT MASA " & Me.cb_pph & " - Tahun & tahun"
    baris = 4
    fLs.Cells(baris, kolom) = "Filter Unit: " & DIVISI & ". Filter KPP: " & kpp
                                    
    baris = 8
    kolom = 2
    fLs.Cells(baris, kolom) = "Data dari awal s/d Bulan:" & Masa_Pajak & " tahun " & _
                            tahun
    
    
    baris = 8
    kolom = 2
    
    Call setListInfo(Me.List1, "Proses Sheet02 - set KPP")
    sql = "Select mkpp.npwp, mkpp.kpp_administrasi From mkpp order by npwp"
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
    End If
    jRec = RecordCount(rs)
    If jRec > 0 Then
        rs.MoveFirst
        c = 1
        Do While rs.EOF = False
            Call info_progress(Me.StatusBar1, 1, c, jRec, "KPP")
            fLs.Cells(baris, 2) = c
            fLs.Cells(baris, 3) = "'" & cek_null(rs(0))
            fLs.Cells(baris, 4) = cek_null(rs(1))
            c = c + 1
            baris = baris + 1
            rs.MoveNext
        Loop
    End If
    
    '-- looping dari bulan 1 s/d x
    kolom = 5
    For bln = 1 To CInt(Masa_Pajak)
        Call setListInfo(Me.List1, "Proses Sheet02 - write Bln" & bln)
        baris = 8
        Set rs = get_pelaporan_spt_masa(DIVISI, tahun, adddigit(CLng(bln), 2), kpp)
        jRec = RecordCount(rs)
        If jRec > 0 Then
            rs.MoveFirst
            c = 1
            Do While rs.EOF = False
                Call info_progress(Me.StatusBar1, 1, c, jRec, "write bln" & bln)
                fLs.Cells(baris, kolom + 0) = cek_null(rs(2))
                'If cek_Money(rs(3)) > 0 Then
                '    delay (1)
                'End If
                fLs.Cells(baris, kolom + 1) = cek_Money(rs(3))  's_Sewa_Tanah_Bangunan_bruto
                fLs.Cells(baris, kolom + 2) = cek_Money(rs(4))
                
                fLs.Cells(baris, kolom + 3) = cek_Money(rs(5))
                fLs.Cells(baris, kolom + 1).NumberFormat = "#,##0"
                fLs.Cells(baris, kolom + 2).NumberFormat = "#,##0"
                fLs.Cells(baris, kolom + 3).NumberFormat = "#,##0"
                fLs.Cells(baris, kolom + 4) = cek_Money(rs(6))       'deviden
                
                fLs.Cells(baris, kolom + 5) = cek_Money(rs(3)) + cek_Money(rs(4)) + _
                                                cek_Money(rs(5)) + cek_Money(rs(6))       'jumlah all
                
                fLs.Cells(baris, kolom + 6) = cek_Money(rs(8))
                fLs.Cells(baris, kolom + 6).NumberFormat = "#,##0"
                
                If cek_Money(rs(3)) > 0 Then
                    fLs.Cells(baris, kolom + 7) = (cek_Money(rs(8)) / cek_Money(rs(3)))
                Else
                    fLs.Cells(baris, kolom + 7) = ""
                End If
                
                fLs.Cells(baris, kolom + 8) = cek_Money(rs(10))
                fLs.Cells(baris, kolom + 8).NumberFormat = "#,##0"
                If cek_Money(rs(4)) > 0 Then
                    fLs.Cells(baris, kolom + 9) = (cek_Money(rs(10)) / cek_Money(rs(4)))
                Else
                    fLs.Cells(baris, kolom + 9) = ""
                End If
                
                fLs.Cells(baris, kolom + 10) = cek_Money(rs(12))
                fLs.Cells(baris, kolom + 11) = cek_null(rs(13))
                fLs.Cells(baris, kolom + 12) = cek_Money(rs(14))
                fLs.Cells(baris, kolom + 13) = cek_null(rs(15))
                fLs.Cells(baris, kolom + 14) = cek_Money(rs(8)) + cek_Money(rs(10)) + _
                                            cek_Money(rs(12)) + cek_Money(rs(14))
                fLs.Cells(baris, kolom + 14).NumberFormat = "#,##0"
                c = c + 1
                baris = baris + 1
                rs.MoveNext
            Loop
        End If
        kolom = kolom + 15
    Next

    
    fl.ActiveWorkbook.Save
    
    'fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub

Sub pph21_sheet_02(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long, bln As Integer
    Dim Total1 As Currency
    
    f = fileSimpan
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet02 - file not found")
            Exit Sub
        End If
    Else
        MsgBox "File tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet02 - file not found")
            Exit Sub
        Exit Sub
    End If
    
    'open sheet 2, isi data looping
    Set fLs = fl.Sheets(2)
    'baris = 2
    'kolom = 2
    
    'fLs.Cells(baris, kolom) = "REKAP PELAPORAN SPT MASA " & Me.cb_pph & " - Tahun & tahun"
    'baris = 4
    'fLs.Cells(baris, kolom) = "Filter Unit: " & DIVISI & ". Filter KPP: " & kpp
                                    
    baris = 1
    kolom = 1
    fLs.Cells(baris, kolom) = "Filter Unit: " & DIVISI & ". Filter KPP: " & kpp & _
                                ".Data dari awal s/d Bulan:" & Masa_Pajak & " tahun " & _
                                tahun
    
    
    baris = 8
    kolom = 2
    
    Call setListInfo(Me.List1, "Proses Sheet02 - set KPP")
    sql = "Select ssp_pph.Tanggal_Setor_SSP, '' as tempatBayar, " & _
        "ssp_pph.NPWP_KPP, mkpp.nama, mkpp.kpp_administrasi, " & _
        "ssp_pph.Kode_KAP, ssp_pph.Jenis_Pajak, ssp_pph.Kode_Jenis_Setoran, " & _
        "ssp_pph.Masa_Pajak, ssp_pph.Tahun_Pajak, ssp_pph.NTPN, " & _
        "mdivisi.nama_divisi, ssp_pph.Jumlah_SSP, '' as ket, " & _
        "ssp_pph.Kode_Form " & _
        "From " & _
        "ssp_pph Left Join " & _
        "mkpp On mkpp.npwp = ssp_pph.NPWP_KPP Left Join " & _
        "mdivisi On mdivisi.kodedivisi = ssp_pph.kode_divisi  " & _
        "where instr(ssp_pph.Jenis_Pajak,'21') > 0 "
    If Trim(kpp) = "" Or Trim(kpp) = "ALL" Then
    Else
        sql = sql & "and ssp_pph.NPWP_KPP = '" & kpp & "' "
    End If
    
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
    Else
        sql = sql & "and ssp_pph.kode_divisi = '" & DIVISI & "' "
    End If
    
    sql = sql & "order by ssp_pph.NPWP_KPP, ssp_pph.Tahun_Pajak, ssp_pph.Masa_Pajak, " & _
            "ssp_pph.kode_divisi"
    
    'sql = InputBox("", "", sql)
        
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
    End If
    jRec = RecordCount(rs)
    If jRec > 0 Then
        rs.MoveFirst
        c = 1
        baris = 3
        Total1 = 0
        Do While rs.EOF = False
            Call info_progress(Me.StatusBar1, 1, c, jRec, "KPP")
            
            For a = 0 To 14
                If a = 12 Then
                    fLs.Cells(baris, a + 1) = cek_Money(rs(a))
                    fLs.Cells(baris, a + 1).NumberFormat = "#,##0"
                    Total1 = Total1 + cek_Money(rs(a))
                Else
                    fLs.Cells(baris, a + 1) = cek_null(rs(a))
                End If
            Next
            c = c + 1
            baris = baris + 1
            rs.MoveNext
        Loop
        
        baris = 1
        kolom = 13
        fLs.Cells(baris, kolom) = cek_Money(Total1)
        fLs.Cells(baris, kolom).NumberFormat = "#,##0"
    Else
        Call setListInfo(Me.List1, "Proses Sheet02 - set KPP - data tidak ditemukan")
    End If
    
    
    
    fl.ActiveWorkbook.Save
    
    'fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub

Sub pph22_sheet_02(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long, bln As Integer
    Dim Total1 As Currency
    
    f = fileSimpan
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet02 - file not found")
            Exit Sub
        End If
    Else
        MsgBox "File tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet02 - file not found")
            Exit Sub
        Exit Sub
    End If
    
    'open sheet 2, isi data looping
    Set fLs = fl.Sheets(2)
    'baris = 2
    'kolom = 2
    
    'fLs.Cells(baris, kolom) = "REKAP PELAPORAN SPT MASA " & Me.cb_pph & " - Tahun & tahun"
    'baris = 4
    'fLs.Cells(baris, kolom) = "Filter Unit: " & DIVISI & ". Filter KPP: " & kpp
                                    
    baris = 1
    kolom = 1
    fLs.Cells(baris, kolom) = "Filter Unit: " & DIVISI & ". Filter KPP: " & kpp & _
                                ".Data dari awal s/d Bulan:" & Masa_Pajak & " tahun " & _
                                tahun
    
    
    baris = 8
    kolom = 2
    
    Call setListInfo(Me.List1, "Proses Sheet02 - set KPP")
    sql = "Select ssp_pph.Tanggal_Setor_SSP, '' as tempatBayar, " & _
        "ssp_pph.NPWP_KPP, mkpp.nama, mkpp.kpp_administrasi, " & _
        "ssp_pph.Kode_KAP, ssp_pph.Jenis_Pajak, ssp_pph.Kode_Jenis_Setoran, " & _
        "ssp_pph.Masa_Pajak, ssp_pph.Tahun_Pajak, ssp_pph.NTPN, " & _
        "mdivisi.nama_divisi, ssp_pph.Jumlah_SSP, '' as ket, " & _
        "ssp_pph.Kode_Form " & _
        "From " & _
        "ssp_pph Left Join " & _
        "mkpp On mkpp.npwp = ssp_pph.NPWP_KPP Left Join " & _
        "mdivisi On mdivisi.kodedivisi = ssp_pph.kode_divisi  " & _
        "where instr(ssp_pph.Jenis_Pajak,'22') > 0 "
    If Trim(kpp) = "" Or Trim(kpp) = "ALL" Then
    Else
        sql = sql & "and ssp_pph.NPWP_KPP = '" & kpp & "' "
    End If
    
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
    Else
        sql = sql & "and ssp_pph.kode_divisi = '" & DIVISI & "' "
    End If
    
    sql = sql & "order by ssp_pph.NPWP_KPP, ssp_pph.Tahun_Pajak, ssp_pph.Masa_Pajak, " & _
            "ssp_pph.kode_divisi"
    
    'sql = InputBox("", "", sql)
        
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
    End If
    jRec = RecordCount(rs)
    If jRec > 0 Then
        rs.MoveFirst
        c = 1
        baris = 3
        Total1 = 0
        Do While rs.EOF = False
            Call info_progress(Me.StatusBar1, 1, c, jRec, "KPP")
            
            For a = 0 To 14
                If a = 12 Then
                    fLs.Cells(baris, a + 1) = cek_Money(rs(a))
                    fLs.Cells(baris, a + 1).NumberFormat = "#,##0"
                    Total1 = Total1 + cek_Money(rs(a))
                Else
                    fLs.Cells(baris, a + 1) = cek_null(rs(a))
                End If
            Next
            c = c + 1
            baris = baris + 1
            rs.MoveNext
        Loop
        
        baris = 1
        kolom = 13
        fLs.Cells(baris, kolom) = cek_Money(Total1)
        fLs.Cells(baris, kolom).NumberFormat = "#,##0"
    Else
        Call setListInfo(Me.List1, "Proses Sheet02 - set KPP - data tidak ditemukan")
    End If
    
    
    
    fl.ActiveWorkbook.Save
    
    'fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub

Sub pph23_sheet_02(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long, bln As Integer
    Dim Total1 As Currency
    
    f = fileSimpan
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet02 - file not found")
            Exit Sub
        End If
    Else
        MsgBox "File tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet02 - file not found")
            Exit Sub
        Exit Sub
    End If
    
    'open sheet 2, isi data looping
    Set fLs = fl.Sheets(2)
    'baris = 2
    'kolom = 2
    
    'fLs.Cells(baris, kolom) = "REKAP PELAPORAN SPT MASA " & Me.cb_pph & " - Tahun & tahun"
    'baris = 4
    'fLs.Cells(baris, kolom) = "Filter Unit: " & DIVISI & ". Filter KPP: " & kpp
                                    
    baris = 1
    kolom = 1
    fLs.Cells(baris, kolom) = "Filter Unit: " & DIVISI & ". Filter KPP: " & kpp & _
                                ".Data dari awal s/d Bulan:" & Masa_Pajak & " tahun " & _
                                tahun
    
    
    baris = 8
    kolom = 2
    
    Call setListInfo(Me.List1, "Proses Sheet02 - set KPP")
    sql = "Select ssp_pph.Tanggal_Setor_SSP, '' as tempatBayar, " & _
        "ssp_pph.NPWP_KPP, mkpp.nama, mkpp.kpp_administrasi, " & _
        "ssp_pph.Kode_KAP, ssp_pph.Jenis_Pajak, ssp_pph.Kode_Jenis_Setoran, " & _
        "ssp_pph.Masa_Pajak, ssp_pph.Tahun_Pajak, ssp_pph.NTPN, " & _
        "mdivisi.nama_divisi, ssp_pph.Jumlah_SSP, '' as ket, " & _
        "ssp_pph.Kode_Form " & _
        "From " & _
        "ssp_pph Left Join " & _
        "mkpp On mkpp.npwp = ssp_pph.NPWP_KPP Left Join " & _
        "mdivisi On mdivisi.kodedivisi = ssp_pph.kode_divisi  " & _
        "where instr(ssp_pph.Jenis_Pajak,'23') > 0 or instr(ssp_pph.Jenis_Pajak,'26') > 0 "
    If Trim(kpp) = "" Or Trim(kpp) = "ALL" Then
    Else
        sql = sql & "and ssp_pph.NPWP_KPP = '" & kpp & "' "
    End If
    
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
    Else
        sql = sql & "and ssp_pph.kode_divisi = '" & DIVISI & "' "
    End If
    
    sql = sql & "order by ssp_pph.NPWP_KPP, ssp_pph.Tahun_Pajak, ssp_pph.Masa_Pajak, " & _
            "ssp_pph.kode_divisi"
    
    'sql = InputBox("", "", sql)
        
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
    End If
    jRec = RecordCount(rs)
    If jRec > 0 Then
        rs.MoveFirst
        c = 1
        baris = 3
        Total1 = 0
        Do While rs.EOF = False
            Call info_progress(Me.StatusBar1, 1, c, jRec, "KPP")
            
            For a = 0 To 14
                If a = 12 Then
                    fLs.Cells(baris, a + 1) = cek_Money(rs(a))
                    fLs.Cells(baris, a + 1).NumberFormat = "#,##0"
                    Total1 = Total1 + cek_Money(rs(a))
                Else
                    fLs.Cells(baris, a + 1) = cek_null(rs(a))
                End If
            Next
            c = c + 1
            baris = baris + 1
            rs.MoveNext
        Loop
        
        baris = 1
        kolom = 13
        fLs.Cells(baris, kolom) = cek_Money(Total1)
        fLs.Cells(baris, kolom).NumberFormat = "#,##0"
    Else
        Call setListInfo(Me.List1, "Proses Sheet02 - set KPP - data tidak ditemukan")
    End If
    
    
    
    fl.ActiveWorkbook.Save
    
    'fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub

Function get_PPh21_spt_masa_noNpwp(DIVISI As String, tahun As String, masa As String) As Recordset
    Dim sql As String
    Dim rs As Recordset
    
    If Trim(UCase(DIVISI)) = "ALL" Then
        DIVISI = ""
    End If
    
    sql = "select 'pegawai tetap' as ket, " & _
        "F_get_kryw_ttp_jml_bulan('" & tahun & "','" & masa & "','" & DIVISI & "') as jml_karyawan, " & _
        "F_get_kryw_ttp_bruto_bulan('" & tahun & "','" & masa & "','" & DIVISI & "') as pbruto,  " & _
        "F_get_kryw_ttp_pph21_bulan('" & tahun & "','" & masa & "','" & DIVISI & "') as pph " & _
        "Union " & _
        "select 'Pegawai Tidak Tetap / T.Kerja Lepas' as ket,  " & _
        "'0' as jml_karyawan, " & _
        "'0' as pbruto, " & _
        "'0' as pph " & _
        "Union " & _
        "select 'pesangon' as ket,  " & _
        "F_get_pesangon_jml_bulan('" & tahun & "','" & masa & "','" & DIVISI & "') as jml_karyawan, " & _
        "F_get_pesangon_bruto_bulan('" & tahun & "','" & masa & "','" & DIVISI & "') as pbruto, " & _
        "F_get_pesangon_pph_bulan('" & tahun & "','" & masa & "','" & DIVISI & "') as pph "
    'sql = InputBox("sql error", "", sql)
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Set get_PPh21_spt_masa_noNpwp = Null
    Else
        Set get_PPh21_spt_masa_noNpwp = rs
    End If
End Function

Function get_PPh22_spt_masa_noNpwp(DIVISI As String, tahun As String, masa As String) As Recordset
    Dim sql As String
    Dim rs As Recordset
    
    If Trim(UCase(DIVISI)) = "ALL" Then
        DIVISI = ""
    End If
    
    sql = "select 'Pembelian Oleh Bendaharawa' as ket, " & _
        "F_get_pembelian_bendahara_bulan('" & tahun & "','" & masa & "','" & DIVISI & "') as pbruto,  " & _
        "F_get_pembelian_bendahara_bulan_pph('" & tahun & "','" & masa & "','" & DIVISI & "') as pph "
    'sql = InputBox("sql error", "", sql)
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Set get_PPh22_spt_masa_noNpwp = Null
    Else
        Set get_PPh22_spt_masa_noNpwp = rs
    End If
End Function

Function get_PPh23_spt_masa_noNpwp(DIVISI As String, tahun As String, masa As String) As Recordset
    Dim sql As String
    Dim rs As Recordset
    
    If Trim(UCase(DIVISI)) = "ALL" Then
        DIVISI = ""
    End If
    
    sql = "select 'Pembelian Oleh Bendahara' as ket, " & _
        "F_get_pembelian_bendahara_bulan('2020','12','') as pbruto, " & _
        "F_get_pembelian_bendahara_bulan_pph('2020','12','') as pph " & _
        "Union " & _
        "select 'Jasa Selain Konstruksi' as ket, " & _
        "F_get_jasa_lain_bulan('2020','12','') as pbruto, " & _
        "F_get_jasa_lain_bulan_pph('2020','12','') as pph " & _
        "Union " & _
        "select 'Deviden' as ket, " & _
        "F_get_deviden_bulan('2020','12','') as pbruto, " & _
        "F_get_deviden_bulan_pph('2020','12','') as pph " & _
        "Union " & _
        "select 'Penggunaan Harta' as ket, " & _
        "F_get_penggunaan_harta_bulan('2020','12','') as pbruto, " & _
        "F_get_penggunaan_harta_bulan_pph('2020','12','') as pph " & _
        "Union " & _
        "select 'Hadiah dan Penghargaan' as ket, " & _
        "F_get_hadiah_bulan('2020','12','') as pbruto, " & _
        "F_get_hadiah_bulan_pph('2020','12','') as pph"
    'sql = InputBox("sql error", "", sql)
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Set get_PPh23_spt_masa_noNpwp = Null
    Else
        Set get_PPh23_spt_masa_noNpwp = rs
    End If
End Function

Function get_PPh42_masa_noNpwp(DIVISI As String, tahun As String, masa As String) As Recordset
    Dim sql As String
    Dim rs As Recordset
    
    If Trim(UCase(DIVISI)) = "ALL" Then
        DIVISI = ""
    End If
    
    sql = "select 'bunga obligasi' as ket, " & _
        "F_get_bungaObligasi_bruto('" & tahun & "','" & masa & "','" & DIVISI & "') as pbruto,  " & _
        "'' as pph " & _
        "Union " & _
        "select 'sewa tanah bangunan' as ket,  " & _
        "F_get_SewaTanahBangunan_bruto('" & tahun & "','" & masa & "','" & DIVISI & "') as pbruto, " & _
        "F_get_SewaTanahBangunan_pph('" & tahun & "','" & masa & "','" & DIVISI & "') as pph " & _
        "Union " & _
        "select 'jasa konstruksi' as ket,  " & _
        "F_get_jasakonstruksi_bruto('" & tahun & "','" & masa & "','" & DIVISI & "') as pbruto, " & _
        "F_get_jasakonstruksi_pph('" & tahun & "','" & masa & "','" & DIVISI & "') as pph "
    'sql = InputBox("sql error", "", sql)
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Set get_PPh42_masa_noNpwp = Null
    Else
        Set get_PPh42_masa_noNpwp = rs
    End If
End Function

Sub pph42_sheet_03(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long, bln As Integer
    Dim Total1 As Currency, totalPBKonstruksi As Currency, totalPBSewa As Currency
    Dim data1, data2, isi
    
    f = fileSimpan
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet03 - file not found")
            Exit Sub
        End If
    Else
        MsgBox "File tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet03 - file not found")
            Exit Sub
        Exit Sub
    End If
    
    'open sheet 3, isi data looping
    Set fLs = fl.Sheets(3)
    baris = 7
    kolom = 4
    
    fLs.Cells(baris, kolom) = ":" & tahun
    baris = 8
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
        fLs.Cells(baris, kolom) = "SPT Masa " & Me.cb_pph & ""
    Else
        fLs.Cells(baris, kolom) = "SPT Masa " & Me.cb_pph & ". Filter Divisi: " & DIVISI
    End If
                                        
    baris = 18
    
    Call setListInfo(Me.List1, "Proses Sheet03 - set SPT Masa")
    '-- looping dari bulan 1 s/d x
    kolom = 5
    For bln = 1 To CInt(12)
        Call setListInfo(Me.List1, "Proses Sheet03 - write Bln" & bln)
        baris = 8
        Set rs = get_PPh42_masa_noNpwp(DIVISI, tahun, adddigit(CLng(bln), 2))
        jRec = RecordCount(rs)
        If jRec > 0 Then
            rs.MoveFirst
            baris = 18
            fLs.Cells(baris, kolom + 0) = cek_Money(rs(1))
            fLs.Cells(baris, kolom + 1) = cek_Money(rs(2))
            fLs.Cells(baris, kolom + 0).NumberFormat = "#,##0"
            fLs.Cells(baris, kolom + 1).NumberFormat = "#,##0"
            
            rs.MoveNext
            baris = 20
            fLs.Cells(baris, kolom + 0) = cek_Money(rs(1))
            fLs.Cells(baris, kolom + 1) = cek_Money(rs(2))
            fLs.Cells(baris, kolom + 0).NumberFormat = "#,##0"
            fLs.Cells(baris, kolom + 1).NumberFormat = "#,##0"
            
            rs.MoveNext
            baris = 21
            fLs.Cells(baris, kolom + 0) = cek_Money(rs(1))
            fLs.Cells(baris, kolom + 1) = cek_Money(rs(2))
            fLs.Cells(baris, kolom + 0).NumberFormat = "#,##0"
            fLs.Cells(baris, kolom + 1).NumberFormat = "#,##0"
            
        End If
        kolom = kolom + 3
    Next

    '--
    kolom = 5
    Call setListInfo(Me.List1, "Proses Sheet03 - saldo akun per bulan")
    For bln = 1 To CInt(12)
        baris = 44
        Call setListInfo(Me.List1, "Proses Sheet03 - write Bln" & bln)
        isi = Array("", get_nilaiAkun_divisi_bln("50101", tahun, CStr(bln), DIVISI), _
                get_nilaiAkun_divisi_bln("50301", tahun, CStr(bln), DIVISI), _
                get_nilaiAkun_divisi_bln("50302", tahun, CStr(bln), DIVISI), _
                get_nilaiAkun_divisi_bln("51851", tahun, CStr(bln), DIVISI), _
                get_nilaiAkun_divisi_bln("83251", tahun, CStr(bln), DIVISI), _
                "", "", _
                get_nilaiAkunAll_divisi("20111", CStr(CInt(tahun) - 1), DIVISI), _
                get_nilaiAkunAll_divisi("20112", CStr(CInt(tahun) - 1), DIVISI), _
                get_nilaiAkunAll_divisi("20113", CStr(CInt(tahun) - 1), DIVISI), _
                get_nilaiAkunAll_divisi("20116", CStr(CInt(tahun) - 1), DIVISI), _
                get_nilaiAkunAll_divisi("20131", CStr(CInt(tahun) - 1), DIVISI), _
                "", "", _
                get_nilaiAkun_divisi_bln("20111", tahun, CStr(bln), DIVISI) * -1, _
                get_nilaiAkun_divisi_bln("20112", tahun, CStr(bln), DIVISI) * -1, _
                get_nilaiAkun_divisi_bln("20113", tahun, CStr(bln), DIVISI) * -1, _
                get_nilaiAkun_divisi_bln("20116", tahun, CStr(bln), DIVISI) * -1, _
                get_nilaiAkun_divisi_bln("20131", tahun, CStr(bln), DIVISI) * -1)
        
        For c = 1 To 19
            If Trim(isi(c)) <> "" Then
                fLs.Cells(baris + c - 1, kolom) = cek_Money(isi(c))
                fLs.Cells(baris + c - 1, kolom).NumberFormat = "#,##0"
            End If
        Next
        kolom = kolom + 3
    Next
    


    fl.ActiveWorkbook.Save
    
    'fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub

Sub pph21_sheet_03(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long, bln As Integer
    Dim Total1 As Currency, totalPBKonstruksi As Currency, totalPBSewa As Currency
    Dim data1, data2, isi
    
    f = fileSimpan
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet03 - file not found")
            Exit Sub
        End If
    Else
        MsgBox "File tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet03 - file not found")
            Exit Sub
        Exit Sub
    End If
    
    'open sheet 3, isi data looping
    Set fLs = fl.Sheets(3)
    baris = 6
    kolom = 4
    
    fLs.Cells(baris, kolom) = tahun
    baris = 7
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
        fLs.Cells(baris, kolom) = "SPT Masa " & Me.cb_pph & ""
    Else
        fLs.Cells(baris, kolom) = "SPT Masa " & Me.cb_pph & ". Filter Divisi: " & DIVISI
    End If
                                        
    baris = 16
    
    Call setListInfo(Me.List1, "Proses Sheet03 - set SPT Masa")
    '-- looping dari bulan 1 s/d x
    kolom = 4
    For bln = 1 To CInt(12)
        Call setListInfo(Me.List1, "Proses Sheet03 - write Bln" & bln)
        baris = 8
        Set rs = get_PPh21_spt_masa_noNpwp(DIVISI, tahun, adddigit(CLng(bln), 2))
        jRec = RecordCount(rs)
        If jRec > 0 Then
            rs.MoveFirst
            baris = 16
            fLs.Cells(baris, kolom + 0) = cek_null(rs(1))
            fLs.Cells(baris, kolom + 1) = cek_Money(rs(2))
            fLs.Cells(baris, kolom + 2) = cek_Money(rs(3))
            fLs.Cells(baris, kolom + 1).NumberFormat = "#,##0"
            fLs.Cells(baris, kolom + 2).NumberFormat = "#,##0"
            
            
            rs.MoveNext
            baris = 18
            fLs.Cells(baris, kolom + 0) = cek_null(rs(1))
            fLs.Cells(baris, kolom + 1) = cek_Money(rs(2))
            fLs.Cells(baris, kolom + 2) = cek_Money(rs(3))
            fLs.Cells(baris, kolom + 1).NumberFormat = "#,##0"
            fLs.Cells(baris, kolom + 2).NumberFormat = "#,##0"
            
            rs.MoveNext
            baris = 32
            fLs.Cells(baris, kolom + 0) = cek_null(rs(1))
            fLs.Cells(baris, kolom + 1) = cek_Money(rs(2))
            fLs.Cells(baris, kolom + 2) = cek_Money(rs(3))
            fLs.Cells(baris, kolom + 1).NumberFormat = "#,##0"
            fLs.Cells(baris, kolom + 2).NumberFormat = "#,##0"
            
        End If
        kolom = kolom + 4
    Next

    '--
    kolom = 5
    Call setListInfo(Me.List1, "Proses Sheet03 - saldo akun per bulan")
    For bln = 1 To CInt(12)
        baris = 45
        Call setListInfo(Me.List1, "Proses Sheet03 - write Bln" & bln)
        isi = Array("", get_nilaiAkun_divisi_bln("51101", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51111", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51114", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51115", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51116", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51119", tahun, CStr(bln), DIVISI), _
                get_nilaiAkun_divisi_bln("51122", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51125", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51201", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51206", tahun, CStr(bln), DIVISI), _
                get_nilaiAkun_divisi_bln("51213", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51215", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51216", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51219", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51221", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51222", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51225", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51228", tahun, CStr(bln), DIVISI), _
                get_nilaiAkun_divisi_bln("51501", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51502", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51861", tahun, CStr(bln), DIVISI), _
                get_nilaiAkun_divisi_bln("80101", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80111", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80114", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80115", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80116", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80118", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80119", tahun, CStr(bln), DIVISI), _
                get_nilaiAkun_divisi_bln("80121", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80123", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80124", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80125", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80126", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80127", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80128", tahun, CStr(bln), DIVISI), _
                get_nilaiAkun_divisi_bln("80131", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80133", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80136", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80137", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80138", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80139", tahun, CStr(bln), DIVISI), _
                get_nilaiAkun_divisi_bln("80201", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80214", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80215", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80216", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80219", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80221", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80223", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80225", tahun, CStr(bln), DIVISI), _
                get_nilaiAkun_divisi_bln("80501", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("80601", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("83261", tahun, CStr(bln), DIVISI), _
                "", "", _
                get_nilaiAkunAll_divisi("20701", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                get_nilaiAkunAll_divisi("20704", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                get_nilaiAkunAll_divisi("20709", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                get_nilaiAkunAll_divisi("21902", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                "", "", _
                get_nilaiAkun_divisi_bln("20701", tahun, CStr(bln), DIVISI) * 1, get_nilaiAkun_divisi_bln("20704", tahun, CStr(bln), DIVISI) * 1, get_nilaiAkun_divisi_bln("20709", tahun, CStr(bln), DIVISI) * 1, get_nilaiAkun_divisi_bln("21902", tahun, CStr(bln), DIVISI) * 1)
        For c = 1 To 64
            If Trim(isi(c)) <> "" Then
                fLs.Cells(baris + c - 1, kolom) = cek_Money(isi(c))
                fLs.Cells(baris + c - 1, kolom).NumberFormat = "#,##0"
            End If
        Next
        kolom = kolom + 4
    Next
    


    fl.ActiveWorkbook.Save
    
    'fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub

Sub pph22_sheet_03(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long, bln As Integer
    Dim Total1 As Currency, totalPBKonstruksi As Currency, totalPBSewa As Currency
    Dim data1, data2, isi
    
    f = fileSimpan
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet03 - file not found")
            Exit Sub
        End If
    Else
        MsgBox "File tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet03 - file not found")
            Exit Sub
        Exit Sub
    End If
    
    'open sheet 3, isi data looping
    Set fLs = fl.Sheets(3)
    baris = 6
    kolom = 4
    
    fLs.Cells(baris, kolom) = tahun
    baris = 7
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
        fLs.Cells(baris, kolom) = "SPT Masa " & Me.cb_pph & ""
    Else
        fLs.Cells(baris, kolom) = "SPT Masa " & Me.cb_pph & ". Filter Divisi: " & DIVISI
    End If
                                        
    baris = 16
    
    Call setListInfo(Me.List1, "Proses Sheet03 - set SPT Masa")
    '-- looping dari bulan 1 s/d x
    kolom = 4
    For bln = 1 To CInt(12)
        Call setListInfo(Me.List1, "Proses Sheet03 - write Bln" & bln)
        baris = 8
        Set rs = get_PPh22_spt_masa_noNpwp(DIVISI, tahun, adddigit(CLng(bln), 2))
        jRec = RecordCount(rs)
        If jRec > 0 Then
            rs.MoveFirst
            baris = 17
            fLs.Cells(baris, kolom + 0) = cek_null(rs(1))
            fLs.Cells(baris, kolom + 1) = cek_Money(rs(2))
            fLs.Cells(baris, kolom + 0).NumberFormat = "#,##0"
            fLs.Cells(baris, kolom + 1).NumberFormat = "#,##0"
            
        End If
        kolom = kolom + 3
    Next

    '--
    kolom = 4
    Call setListInfo(Me.List1, "Proses Sheet03 - saldo akun per bulan")
    For bln = 1 To CInt(12)
        baris = 36
        Call setListInfo(Me.List1, "Proses Sheet03 - write Bln" & bln)
        isi = Array("", get_nilaiAkun_divisi_bln("50201", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("50202", tahun, CStr(bln), DIVISI), _
                "", "", _
                get_nilaiAkunAll_divisi("20101", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                get_nilaiAkunAll_divisi("20102", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                "", "", _
                get_nilaiAkun_divisi_bln("20101", tahun, CStr(bln), DIVISI) * 1, _
                get_nilaiAkun_divisi_bln("20102", tahun, CStr(bln), DIVISI) * 1)
        For c = 1 To 10
            If Trim(isi(c)) <> "" Then
                fLs.Cells(baris + c - 1, kolom) = cek_Money(isi(c))
                fLs.Cells(baris + c - 1, kolom).NumberFormat = "#,##0"
            End If
        Next
        kolom = kolom + 3
    Next
    


    fl.ActiveWorkbook.Save
    
    'fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub

Sub pph23_sheet_03(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long, bln As Integer
    Dim Total1 As Currency, totalPBKonstruksi As Currency, totalPBSewa As Currency
    Dim data1, data2, isi
    
    f = fileSimpan
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet03 - file not found")
            Exit Sub
        End If
    Else
        MsgBox "File tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet03 - file not found")
            Exit Sub
        Exit Sub
    End If
    
    'open sheet 3, isi data looping
    Set fLs = fl.Sheets(3)
    baris = 6
    kolom = 4
    
    fLs.Cells(baris, kolom) = tahun
    baris = 7
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
        fLs.Cells(baris, kolom) = "SPT Masa " & Me.cb_pph & ""
    Else
        fLs.Cells(baris, kolom) = "SPT Masa " & Me.cb_pph & ". Filter Divisi: " & DIVISI
    End If
                                        
    baris = 16
    
    Call setListInfo(Me.List1, "Proses Sheet03 - set SPT Masa")
    '-- looping dari bulan 1 s/d x
    kolom = 4
    For bln = 1 To CInt(12)
        Call setListInfo(Me.List1, "Proses Sheet03 - write Bln" & bln)
        baris = 8
        Set rs = get_PPh23_spt_masa_noNpwp(DIVISI, tahun, adddigit(CLng(bln), 2))
        jRec = RecordCount(rs)
        If jRec > 0 Then
            rs.MoveFirst
            baris = 19
            fLs.Cells(baris, kolom + 0) = cek_null(rs(1))
            fLs.Cells(baris, kolom + 1) = cek_Money(rs(2))
            fLs.Cells(baris, kolom + 0).NumberFormat = "#,##0"
            fLs.Cells(baris, kolom + 1).NumberFormat = "#,##0"
            
            rs.MoveNext
            baris = 20
            fLs.Cells(baris, kolom + 0) = cek_null(rs(1))
            fLs.Cells(baris, kolom + 1) = cek_Money(rs(2))
            fLs.Cells(baris, kolom + 0).NumberFormat = "#,##0"
            fLs.Cells(baris, kolom + 1).NumberFormat = "#,##0"
            
            rs.MoveNext
            baris = 23
            fLs.Cells(baris, kolom + 0) = cek_null(rs(1))
            fLs.Cells(baris, kolom + 1) = cek_Money(rs(2))
            fLs.Cells(baris, kolom + 0).NumberFormat = "#,##0"
            fLs.Cells(baris, kolom + 1).NumberFormat = "#,##0"
            
            rs.MoveNext
            baris = 25
            fLs.Cells(baris, kolom + 0) = cek_null(rs(1))
            fLs.Cells(baris, kolom + 1) = cek_Money(rs(2))
            fLs.Cells(baris, kolom + 0).NumberFormat = "#,##0"
            fLs.Cells(baris, kolom + 1).NumberFormat = "#,##0"
            
            rs.MoveNext
            baris = 27
            fLs.Cells(baris, kolom + 0) = cek_null(rs(1))
            fLs.Cells(baris, kolom + 1) = cek_Money(rs(2))
            fLs.Cells(baris, kolom + 0).NumberFormat = "#,##0"
            fLs.Cells(baris, kolom + 1).NumberFormat = "#,##0"
            
        End If
        kolom = kolom + 3
    Next

    '--
    kolom = 4
    Call setListInfo(Me.List1, "Proses Sheet03 - saldo akun per bulan")
    For bln = 1 To CInt(12)
        baris = 37
        Call setListInfo(Me.List1, "Proses Sheet03 - write Bln" & bln)
        isi = Array("", get_nilaiAkun_divisi_bln("50103", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("50401", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("50402", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("50405", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("50406", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("50411", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("50412", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("50413", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("50414", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("50431", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("50432", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("51401", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51402", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("51404", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("51801", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51802", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("51803", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51852", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("51853", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("51901", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("51903", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("71101", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("71102", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("71103", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("81104", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("81121", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("81122", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("83107", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("83211", tahun, CStr(bln), DIVISI), _
         get_nilaiAkun_divisi_bln("83212", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("83252", tahun, CStr(bln), DIVISI), get_nilaiAkun_divisi_bln("83253", tahun, CStr(bln), DIVISI), _
                "", "", _
                get_nilaiAkunAll_divisi("20133", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                get_nilaiAkunAll_divisi("20138", CStr(CInt(tahun) - 1), DIVISI) * -1, _
                "", "", _
                get_nilaiAkun_divisi_bln("20133", tahun, CStr(bln), DIVISI) * 1, _
                get_nilaiAkun_divisi_bln("20138", tahun, CStr(bln), DIVISI) * 1)
        For c = 1 To 40
            If Trim(isi(c)) <> "" Then
                fLs.Cells(baris + c - 1, kolom) = cek_Money(isi(c))
                fLs.Cells(baris + c - 1, kolom).NumberFormat = "#,##0"
            End If
        Next
        kolom = kolom + 3
    Next
    


    fl.ActiveWorkbook.Save
    
    'fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub

Function get_data_akuntansi_sewa(tahun As String, DIVISI As String) As Currency
    Dim sql As String, t As String
    Dim a51851 As Currency, a83251 As Currency
    
    'akun 51851 + 83251
    
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
        DIVISI = ""
    End If
    
    sql = "select F_get_nilaiAkunAll_divisi('51851','" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    a51851 = cek_Money(t)
    
    sql = "select F_get_nilaiAkunAll_divisi('83251','" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    a83251 = cek_Money(t)
    
    get_data_akuntansi_sewa = a51851 + a83251
End Function

Function get_data_akuntansi_PPh21(tahun As String, DIVISI As String) As Currency
    Dim Total1 As Currency
    
    Total1 = get_nilaiAkunAll_divisi("51101", tahun, DIVISI) + get_nilaiAkunAll_divisi("51111", tahun, DIVISI) + get_nilaiAkunAll_divisi("51114", tahun, DIVISI) + get_nilaiAkunAll_divisi("51115", tahun, DIVISI) + get_nilaiAkunAll_divisi("51116", tahun, DIVISI) + get_nilaiAkunAll_divisi("51119", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("51122", tahun, DIVISI) + get_nilaiAkunAll_divisi("51125", tahun, DIVISI) + get_nilaiAkunAll_divisi("51201", tahun, DIVISI) + get_nilaiAkunAll_divisi("51206", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("51213", tahun, DIVISI) + get_nilaiAkunAll_divisi("51215", tahun, DIVISI) + get_nilaiAkunAll_divisi("51216", tahun, DIVISI) + get_nilaiAkunAll_divisi("51219", tahun, DIVISI) + get_nilaiAkunAll_divisi("51221", tahun, DIVISI) + get_nilaiAkunAll_divisi("51222", tahun, DIVISI) + get_nilaiAkunAll_divisi("51225", tahun, DIVISI) + get_nilaiAkunAll_divisi("51228", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("51501", tahun, DIVISI) + get_nilaiAkunAll_divisi("51502", tahun, DIVISI) + get_nilaiAkunAll_divisi("51861", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("80101", tahun, DIVISI) + get_nilaiAkunAll_divisi("80111", tahun, DIVISI) + get_nilaiAkunAll_divisi("80114", tahun, DIVISI) + get_nilaiAkunAll_divisi("80115", tahun, DIVISI) + get_nilaiAkunAll_divisi("80116", tahun, DIVISI) + get_nilaiAkunAll_divisi("80118", tahun, DIVISI) + get_nilaiAkunAll_divisi("80119", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("80121", tahun, DIVISI) + get_nilaiAkunAll_divisi("80123", tahun, DIVISI) + get_nilaiAkunAll_divisi("80124", tahun, DIVISI) + get_nilaiAkunAll_divisi("80125", tahun, DIVISI) + get_nilaiAkunAll_divisi("80126", tahun, DIVISI) + get_nilaiAkunAll_divisi("80127", tahun, DIVISI) + get_nilaiAkunAll_divisi("80128", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("80131", tahun, DIVISI) + get_nilaiAkunAll_divisi("80133", tahun, DIVISI) + get_nilaiAkunAll_divisi("80136", tahun, DIVISI) + get_nilaiAkunAll_divisi("80137", tahun, DIVISI) + get_nilaiAkunAll_divisi("80138", tahun, DIVISI) + get_nilaiAkunAll_divisi("80139", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("80201", tahun, DIVISI) + get_nilaiAkunAll_divisi("80214", tahun, DIVISI) + get_nilaiAkunAll_divisi("80215", tahun, DIVISI) + get_nilaiAkunAll_divisi("80216", tahun, DIVISI) + get_nilaiAkunAll_divisi("80219", tahun, DIVISI) + get_nilaiAkunAll_divisi("80221", tahun, DIVISI) + get_nilaiAkunAll_divisi("80223", tahun, DIVISI) + get_nilaiAkunAll_divisi("80225", tahun, DIVISI) + _
                get_nilaiAkunAll_divisi("80501", tahun, DIVISI) + get_nilaiAkunAll_divisi("80601", tahun, DIVISI) + get_nilaiAkunAll_divisi("83261", tahun, DIVISI) + _
                (get_nilaiAkunAll_divisi("20701", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20704", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20709", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("21902", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20701", tahun, DIVISI) * 1) + _
                (get_nilaiAkunAll_divisi("20704", tahun, DIVISI) * 1) + _
                (get_nilaiAkunAll_divisi("20709", tahun, DIVISI) * 1) + _
                (get_nilaiAkunAll_divisi("21902", tahun, DIVISI) * 1)
    
    get_data_akuntansi_PPh21 = Total1
End Function

Function get_data_akuntansi_PPh22(tahun As String, DIVISI As String) As Currency
    Dim Total1 As Currency
    
    Total1 = get_nilaiAkunAll_divisi("50201", tahun, DIVISI) + get_nilaiAkunAll_divisi("50202", tahun, DIVISI) + _
                (get_nilaiAkunAll_divisi("20101", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20102", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20101", tahun, DIVISI) * 1) + _
                (get_nilaiAkunAll_divisi("20102", tahun, DIVISI) * 1)
    
    get_data_akuntansi_PPh22 = Total1
End Function

Function get_data_akuntansi_PPh23(tahun As String, DIVISI As String) As Currency
    Dim Total1 As Currency
    
    Total1 = get_nilaiAkunAll_divisi("50103", tahun, DIVISI) + get_nilaiAkunAll_divisi("50401", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("50402", tahun, DIVISI) + get_nilaiAkunAll_divisi("50405", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("50406", tahun, DIVISI) + get_nilaiAkunAll_divisi("50411", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("50412", tahun, DIVISI) + get_nilaiAkunAll_divisi("50413", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("50414", tahun, DIVISI) + get_nilaiAkunAll_divisi("50431", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("50432", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("51401", tahun, DIVISI) + get_nilaiAkunAll_divisi("51402", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("51404", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("51801", tahun, DIVISI) + get_nilaiAkunAll_divisi("51802", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("51803", tahun, DIVISI) + get_nilaiAkunAll_divisi("51852", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("51853", tahun, DIVISI) + get_nilaiAkunAll_divisi("51901", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("51903", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("71101", tahun, DIVISI) + get_nilaiAkunAll_divisi("71102", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("71103", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("81104", tahun, DIVISI) + get_nilaiAkunAll_divisi("81121", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("81122", tahun, DIVISI) + get_nilaiAkunAll_divisi("83107", tahun, DIVISI) + get_nilaiAkunAll_divisi("83211", tahun, DIVISI) + _
         get_nilaiAkunAll_divisi("83212", tahun, DIVISI) + get_nilaiAkunAll_divisi("83252", tahun, DIVISI) + get_nilaiAkunAll_divisi("83253", tahun, DIVISI) + _
                (get_nilaiAkunAll_divisi("20133", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20138", CStr(CInt(tahun) - 1), DIVISI) * -1) + _
                (get_nilaiAkunAll_divisi("20133", tahun, DIVISI) * 1) + _
                (get_nilaiAkunAll_divisi("20138", tahun, DIVISI) * 1)
    
    get_data_akuntansi_PPh23 = Total1
End Function

Function get_data_akuntansi_konstruksi(tahun As String, DIVISI As String) As Currency
    Dim sql As String, t As String
    Dim a50101 As Currency, a50301 As Currency
    Dim hutAwal_20111 As Currency, hutAwal_20112 As Currency
    Dim hutAwal_20113 As Currency, hutAwal_20116 As Currency
    Dim hutAwal_20131 As Currency
    
    Dim hutAkhir_20111 As Currency, hutAkhir_20112 As Currency
    Dim hutAkhir_20113 As Currency, hutAkhir_20116 As Currency
    Dim hutAkhir_20131 As Currency
    
    
    '50101 + 50301
    'hutang awal:
    '20111 + 20112 + 20113 + 20116 + 20131
    'hutang akhir
    '20111 + 20112 + 20113 + 20116 + 20131
    
    'If divisi = "300000" Then
    '    delay (1)
    'End If
    
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
        DIVISI = ""
    End If
    
    sql = "select F_get_nilaiAkunAll_divisi('50101','" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    a50101 = cek_Money(t)
    
    sql = "select F_get_nilaiAkunAll_divisi('50301','" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    a50301 = cek_Money(t)
    
    '-----
    sql = "select F_get_nilaiAkunAll_divisi('20111','" & CInt(tahun - 1) & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    hutAwal_20111 = cek_Money(t) * -1
    
    sql = "select F_get_nilaiAkunAll_divisi('20112','" & CInt(tahun - 1) & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    hutAwal_20112 = cek_Money(t) * -1
    
    sql = "select F_get_nilaiAkunAll_divisi('20113','" & CInt(tahun - 1) & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    hutAwal_20113 = cek_Money(t) * -1
    
    sql = "select F_get_nilaiAkunAll_divisi('20116','" & CInt(tahun - 1) & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    hutAwal_20116 = cek_Money(t) * -1
    
    sql = "select F_get_nilaiAkunAll_divisi('20131','" & CInt(tahun - 1) & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    hutAwal_20131 = cek_Money(t) * -1
    
    '-------
    sql = "select F_get_nilaiAkunAll_divisi('20111','" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    hutAkhir_20111 = cek_Money(t)
    
    sql = "select F_get_nilaiAkunAll_divisi('20112','" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    hutAkhir_20112 = cek_Money(t)
    
    sql = "select F_get_nilaiAkunAll_divisi('20113','" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    hutAkhir_20113 = cek_Money(t)
    
    sql = "select F_get_nilaiAkunAll_divisi('20116','" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    hutAkhir_20116 = cek_Money(t)
    
    sql = "select F_get_nilaiAkunAll_divisi('20131','" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    hutAkhir_20131 = cek_Money(t)
    
    '------
    
    get_data_akuntansi_konstruksi = a50101 + a50301 + _
            hutAwal_20111 + hutAwal_20112 + hutAwal_20113 + hutAwal_20116 + hutAwal_20131 + _
            hutAkhir_20111 + hutAkhir_20112 + hutAkhir_20113 + hutAkhir_20116 + hutAkhir_20131
End Function

Function get_data_SPTSewa(tahun As String, DIVISI As String) As Currency
    Dim sql As String, t As String
    
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
        DIVISI = ""
    End If
    sql = "select F_get_pph42_sewa_bruto_thn('" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    get_data_SPTSewa = cek_Money(t)
End Function

Function get_data_SPTPPh21(tahun As String, DIVISI As String) As Currency
    Dim sql As String, t As String
    Dim pph21_bruto_thn As Currency, pph21_pesangon_bruto_thn As Currency
    
    
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
        DIVISI = ""
    End If
    sql = "select F_get_pph21_bruto_thn('" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    pph21_bruto_thn = cek_Money(t)
    
    sql = "select F_get_pph21_pesangon_bruto_thn('" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    pph21_pesangon_bruto_thn = cek_Money(t)
    
    get_data_SPTPPh21 = pph21_bruto_thn + pph21_pesangon_bruto_thn
    
End Function

Function get_data_SPTPPh22(tahun As String, DIVISI As String) As Currency
    Dim sql As String, t As String
    
    
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
        DIVISI = ""
    End If
    sql = "select F_get_pph22_bruto_thn('" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    get_data_SPTPPh22 = cek_Money(t)
    
End Function

Function get_data_SPTPPh23(tahun As String, DIVISI As String) As Currency
    Dim sql As String, t As String
    Dim bruto_pph23 As Currency, bruto_pph26 As Currency
    
    
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
        DIVISI = ""
    End If
    sql = "select F_get_pph23_bruto_thn('" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    bruto_pph23 = cek_Money(t)
    
    sql = "select F_get_pph26_bruto_thn('" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    bruto_pph26 = cek_Money(t)
    
    get_data_SPTPPh23 = bruto_pph23 + bruto_pph26
    
End Function

Function get_data_SPTKonstruksi(tahun As String, DIVISI As String) As Currency
    Dim sql As String, t As String
    
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
        DIVISI = ""
    End If
    sql = "select F_get_pph42_konstruksi_bruto_thn('" & tahun & "','" & DIVISI & "')"
    t = cari_data1(cnn, sql, True)
    get_data_SPTKonstruksi = cek_Money(t)
End Function

Function get_PPh_SPT_ssp(tahun As String, DIVISI As String, jenis As String) As Currency
    Dim sql As String, t As String
    
    If Trim(DIVISI) = "" Or Trim(DIVISI) = "ALL" Then
        DIVISI = ""
    End If
    sql = "select F_get_ssp_pph_thn('" & tahun & "','" & DIVISI & "', '" & jenis & "')"
    t = cari_data1(cnn, sql, True)
    get_PPh_SPT_ssp = cek_Money(t)

End Function

Sub pph42_sheet_04(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer, nmbulan
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long, bln As Integer
    Dim Total1 As Currency, totalPBKonstruksi As Currency, totalPBSewa As Currency
    Dim data1, data2, isi
    Dim DtAkuntansiSewa As Currency, DtSPTSewa As Currency, DtSPTSewa_PPh  As Currency
    Dim DtAkuntansiKonstruksi As Currency, DtSPTKonstruksi As Currency
    
    f = fileSimpan
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet04 - file not found")
            Exit Sub
        End If
    Else
        MsgBox "File tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet04 - file not found")
            Exit Sub
        Exit Sub
    End If
    
    'open sheet 4, isi data looping
    Set fLs = fl.Sheets(4)
    baris = 3
    kolom = 2
    
    fLs.Cells(baris, kolom) = Me.cb_pph & " - Tahun " & tahun
    Call setListInfo(Me.List1, "Proses Sheet04 - Load Laporan")
    
    nmbulan = Array("", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", _
                    "Agustus", "September", "Oktober", "November", "Desember")
    
    'sewa
    'gedung1 / 400000 && gedung2 / 500000
    baris = 6
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiSewa = get_data_akuntansi_sewa(tahun, "400000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiSewa
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTSewa = get_data_SPTSewa(tahun, "400000")
    fLs.Cells(baris + 3, 3) = DtSPTSewa
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTSewa_PPh = get_PPh_SPT_ssp(tahun, "400000", "PPH FINAL SEWA")
    fLs.Cells(baris + 3, 4) = DtSPTSewa_PPh
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    DtAkuntansiSewa = get_data_akuntansi_sewa(tahun, "500000")
    fLs.Cells(baris + 2, 8) = DtAkuntansiSewa
    fLs.Cells(baris + 2, 8).NumberFormat = "#,##0"
    DtSPTSewa = get_data_SPTSewa(tahun, "500000")
    fLs.Cells(baris + 3, 8) = DtSPTSewa
    fLs.Cells(baris + 3, 8).NumberFormat = "#,##0"
    DtSPTSewa_PPh = get_PPh_SPT_ssp(tahun, "500000", "PPH FINAL SEWA")
    fLs.Cells(baris + 3, 9) = DtSPTSewa_PPh
    fLs.Cells(baris + 3, 9).NumberFormat = "#,##0"
    
    
    'infra1 / 200000 && infra2 / 300000
    baris = 14
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiSewa = get_data_akuntansi_sewa(tahun, "200000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiSewa
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTSewa = get_data_SPTSewa(tahun, "200000")
    fLs.Cells(baris + 3, 3) = DtSPTSewa
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTSewa_PPh = get_PPh_SPT_ssp(tahun, "200000", "PPH FINAL SEWA")
    fLs.Cells(baris + 3, 4) = DtSPTSewa_PPh
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    DtAkuntansiSewa = get_data_akuntansi_sewa(tahun, "300000")
    fLs.Cells(baris + 2, 8) = DtAkuntansiSewa
    fLs.Cells(baris + 2, 8).NumberFormat = "#,##0"
    DtSPTSewa = get_data_SPTSewa(tahun, "300000")
    fLs.Cells(baris + 3, 8) = DtSPTSewa
    fLs.Cells(baris + 3, 8).NumberFormat = "#,##0"
    DtSPTSewa_PPh = get_PPh_SPT_ssp(tahun, "300000", "PPH FINAL SEWA")
    fLs.Cells(baris + 3, 9) = DtSPTSewa_PPh
    fLs.Cells(baris + 3, 9).NumberFormat = "#,##0"
    
    'EPC / 780000
    baris = 22
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiSewa = get_data_akuntansi_sewa(tahun, "780000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiSewa
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTSewa = get_data_SPTSewa(tahun, "780000")
    fLs.Cells(baris + 3, 3) = DtSPTSewa
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTSewa_PPh = get_PPh_SPT_ssp(tahun, "780000", "PPH FINAL SEWA")
    fLs.Cells(baris + 3, 4) = DtSPTSewa_PPh
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    
    'UKP / 100000
    baris = 30
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiSewa = get_data_akuntansi_sewa(tahun, "780000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiSewa
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTSewa = get_data_SPTSewa(tahun, "780000")
    fLs.Cells(baris + 3, 3) = DtSPTSewa
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTSewa_PPh = get_PPh_SPT_ssp(tahun, "780000", "PPH FINAL SEWA")
    fLs.Cells(baris + 3, 4) = DtSPTSewa_PPh
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    '== konstruksi======================
    
    'gedung1 / 400000 && gedung2 / 500000
    baris = 42
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiKonstruksi = get_data_akuntansi_konstruksi(tahun, "400000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiKonstruksi
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTKonstruksi = get_data_SPTKonstruksi(tahun, "400000")
    fLs.Cells(baris + 3, 3) = DtSPTKonstruksi
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTSewa_PPh = get_PPh_SPT_ssp(tahun, "400000", "PPH FINAL")
    fLs.Cells(baris + 3, 4) = DtSPTSewa_PPh
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    DtAkuntansiKonstruksi = get_data_akuntansi_konstruksi(tahun, "500000")
    fLs.Cells(baris + 2, 8) = DtAkuntansiKonstruksi
    fLs.Cells(baris + 2, 8).NumberFormat = "#,##0"
    DtSPTKonstruksi = get_data_SPTKonstruksi(tahun, "500000")
    fLs.Cells(baris + 3, 8) = DtSPTKonstruksi
    fLs.Cells(baris + 3, 8).NumberFormat = "#,##0"
    DtSPTSewa_PPh = get_PPh_SPT_ssp(tahun, "500000", "PPH FINAL")
    fLs.Cells(baris + 3, 9) = DtSPTSewa_PPh
    fLs.Cells(baris + 3, 9).NumberFormat = "#,##0"
    
    
    'infra1 / 200000 && infra2 / 300000
    baris = 50
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiKonstruksi = get_data_akuntansi_konstruksi(tahun, "200000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiKonstruksi
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTKonstruksi = get_data_SPTKonstruksi(tahun, "200000")
    fLs.Cells(baris + 3, 3) = DtSPTKonstruksi
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTSewa_PPh = get_PPh_SPT_ssp(tahun, "200000", "PPH FINAL")
    fLs.Cells(baris + 3, 4) = DtSPTSewa_PPh
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    DtAkuntansiKonstruksi = get_data_akuntansi_konstruksi(tahun, "300000")
    fLs.Cells(baris + 2, 8) = DtAkuntansiKonstruksi
    fLs.Cells(baris + 2, 8).NumberFormat = "#,##0"
    DtSPTKonstruksi = get_data_SPTKonstruksi(tahun, "300000")
    fLs.Cells(baris + 3, 8) = DtSPTKonstruksi
    fLs.Cells(baris + 3, 8).NumberFormat = "#,##0"
    DtSPTSewa_PPh = get_PPh_SPT_ssp(tahun, "300000", "PPH FINAL")
    fLs.Cells(baris + 3, 9) = DtSPTSewa_PPh
    fLs.Cells(baris + 3, 9).NumberFormat = "#,##0"
    
    'EPC / 780000
    baris = 58
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiKonstruksi = get_data_akuntansi_konstruksi(tahun, "780000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiKonstruksi
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTKonstruksi = get_data_SPTKonstruksi(tahun, "780000")
    fLs.Cells(baris + 3, 3) = DtSPTKonstruksi
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTSewa_PPh = get_PPh_SPT_ssp(tahun, "780000", "PPH FINAL")
    fLs.Cells(baris + 3, 4) = DtSPTSewa_PPh
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    
    'UKP / 100000
    baris = 66
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiKonstruksi = get_data_akuntansi_konstruksi(tahun, "780000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiKonstruksi
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTKonstruksi = get_data_SPTKonstruksi(tahun, "780000")
    fLs.Cells(baris + 3, 3) = DtSPTKonstruksi
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTSewa_PPh = get_PPh_SPT_ssp(tahun, "780000", "PPH FINAL")
    fLs.Cells(baris + 3, 4) = DtSPTSewa_PPh
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"

    fl.ActiveWorkbook.Save
    
    'fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub

Sub pph21_sheet_04(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer, nmbulan
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long, bln As Integer
    Dim Total1 As Currency, totalPBKonstruksi As Currency, totalPBSewa As Currency
    Dim data1, data2, isi
    Dim DtAkuntansiPPh21 As Currency, DtSPTPPh21 As Currency, DtSPTPPh21_pph  As Currency
    
    f = fileSimpan
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet04 - file not found")
            Exit Sub
        End If
    Else
        MsgBox "File tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet04 - file not found")
            Exit Sub
        Exit Sub
    End If
    
    'open sheet 4, isi data looping
    Set fLs = fl.Sheets(4)
    baris = 3
    kolom = 2
    
    fLs.Cells(baris, kolom) = Me.cb_pph & " - Tahun " & tahun
    Call setListInfo(Me.List1, "Proses Sheet04 - Load Laporan")
    
    nmbulan = Array("", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", _
                    "Agustus", "September", "Oktober", "November", "Desember")
    
    'sewa
    'gedung1 / 400000 && gedung2 / 500000
    baris = 3
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiPPh21 = get_data_akuntansi_PPh21(tahun, "400000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiPPh21
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTPPh21 = get_data_SPTPPh21(tahun, "400000")
    fLs.Cells(baris + 3, 3) = DtSPTPPh21
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTPPh21_pph = get_PPh_SPT_ssp(tahun, "400000", "21")
    fLs.Cells(baris + 3, 4) = DtSPTPPh21_pph
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    DtAkuntansiPPh21 = get_data_akuntansi_PPh21(tahun, "500000")
    fLs.Cells(baris + 2, 8) = DtAkuntansiPPh21
    fLs.Cells(baris + 2, 8).NumberFormat = "#,##0"
    DtSPTPPh21 = get_data_SPTPPh21(tahun, "500000")
    fLs.Cells(baris + 3, 8) = DtSPTPPh21
    fLs.Cells(baris + 3, 8).NumberFormat = "#,##0"
    DtSPTPPh21_pph = get_PPh_SPT_ssp(tahun, "500000", "21")
    fLs.Cells(baris + 3, 9) = DtSPTPPh21_pph
    fLs.Cells(baris + 3, 9).NumberFormat = "#,##0"
    
    
    'infra1 / 200000 && infra2 / 300000
    baris = 14
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiPPh21 = get_data_akuntansi_PPh21(tahun, "200000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiPPh21
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTPPh21 = get_data_SPTPPh21(tahun, "200000")
    fLs.Cells(baris + 3, 3) = DtSPTPPh21
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTPPh21_pph = get_PPh_SPT_ssp(tahun, "200000", "21")
    fLs.Cells(baris + 3, 4) = DtSPTPPh21_pph
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    DtAkuntansiPPh21 = get_data_akuntansi_PPh21(tahun, "300000")
    fLs.Cells(baris + 2, 8) = DtAkuntansiPPh21
    fLs.Cells(baris + 2, 8).NumberFormat = "#,##0"
    DtSPTPPh21 = get_data_SPTPPh21(tahun, "300000")
    fLs.Cells(baris + 3, 8) = DtSPTPPh21
    fLs.Cells(baris + 3, 8).NumberFormat = "#,##0"
    DtSPTPPh21_pph = get_PPh_SPT_ssp(tahun, "300000", "21")
    fLs.Cells(baris + 3, 9) = DtSPTPPh21_pph
    fLs.Cells(baris + 3, 9).NumberFormat = "#,##0"
    
    'EPC / 780000
    baris = 22
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiPPh21 = get_data_akuntansi_PPh21(tahun, "780000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiPPh21
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTPPh21 = get_data_SPTPPh21(tahun, "780000")
    fLs.Cells(baris + 3, 3) = DtSPTPPh21
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTPPh21_pph = get_PPh_SPT_ssp(tahun, "780000", "21")
    fLs.Cells(baris + 3, 4) = DtSPTPPh21_pph
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    
    'UKP / 100000
    baris = 30
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiPPh21 = get_data_akuntansi_PPh21(tahun, "780000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiPPh21
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTPPh21 = get_data_SPTPPh21(tahun, "780000")
    fLs.Cells(baris + 3, 3) = DtSPTPPh21
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTPPh21_pph = get_PPh_SPT_ssp(tahun, "780000", "21")
    fLs.Cells(baris + 3, 4) = DtSPTPPh21_pph
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    


    fl.ActiveWorkbook.Save
    
    'fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub

Sub pph22_sheet_04(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer, nmbulan
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long, bln As Integer
    Dim Total1 As Currency, totalPBKonstruksi As Currency, totalPBSewa As Currency
    Dim data1, data2, isi
    Dim DtAkuntansiPPh22 As Currency, DtSPTPPh22 As Currency, DtSPTPPh22_pph  As Currency
    
    f = fileSimpan
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet04 - file not found")
            Exit Sub
        End If
    Else
        MsgBox "File tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet04 - file not found")
            Exit Sub
        Exit Sub
    End If
    
    'open sheet 4, isi data looping
    Set fLs = fl.Sheets(4)
    baris = 3
    kolom = 2
    
    fLs.Cells(baris, kolom) = Me.cb_pph & " - Tahun " & tahun
    Call setListInfo(Me.List1, "Proses Sheet04 - Load Laporan")
    
    nmbulan = Array("", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", _
                    "Agustus", "September", "Oktober", "November", "Desember")
    
    'sewa
    'gedung1 / 400000 && gedung2 / 500000
    baris = 3
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiPPh22 = get_data_akuntansi_PPh22(tahun, "400000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiPPh22
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTPPh22 = get_data_SPTPPh22(tahun, "400000")
    fLs.Cells(baris + 3, 3) = DtSPTPPh22
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTPPh22_pph = get_PPh_SPT_ssp(tahun, "400000", "22")
    fLs.Cells(baris + 3, 4) = DtSPTPPh22_pph
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    DtAkuntansiPPh22 = get_data_akuntansi_PPh22(tahun, "500000")
    fLs.Cells(baris + 2, 8) = DtAkuntansiPPh22
    fLs.Cells(baris + 2, 8).NumberFormat = "#,##0"
    DtSPTPPh22 = get_data_SPTPPh21(tahun, "500000")
    fLs.Cells(baris + 3, 8) = DtSPTPPh22
    fLs.Cells(baris + 3, 8).NumberFormat = "#,##0"
    DtSPTPPh22_pph = get_PPh_SPT_ssp(tahun, "500000", "22")
    fLs.Cells(baris + 3, 9) = DtSPTPPh22_pph
    fLs.Cells(baris + 3, 9).NumberFormat = "#,##0"
    
    
    'infra1 / 200000 && infra2 / 300000
    baris = 11
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiPPh22 = get_data_akuntansi_PPh22(tahun, "200000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiPPh22
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTPPh22 = get_data_SPTPPh22(tahun, "200000")
    fLs.Cells(baris + 3, 3) = DtSPTPPh22
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTPPh22_pph = get_PPh_SPT_ssp(tahun, "200000", "22")
    fLs.Cells(baris + 3, 4) = DtSPTPPh22_pph
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    DtAkuntansiPPh22 = get_data_akuntansi_PPh22(tahun, "300000")
    fLs.Cells(baris + 2, 8) = DtAkuntansiPPh22
    fLs.Cells(baris + 2, 8).NumberFormat = "#,##0"
    DtSPTPPh22 = get_data_SPTPPh22(tahun, "300000")
    fLs.Cells(baris + 3, 8) = DtSPTPPh22
    fLs.Cells(baris + 3, 8).NumberFormat = "#,##0"
    DtSPTPPh22_pph = get_PPh_SPT_ssp(tahun, "300000", "22")
    fLs.Cells(baris + 3, 9) = DtSPTPPh22_pph
    fLs.Cells(baris + 3, 9).NumberFormat = "#,##0"
    
    'EPC / 780000
    baris = 19
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiPPh22 = get_data_akuntansi_PPh22(tahun, "780000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiPPh22
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTPPh22 = get_data_SPTPPh22(tahun, "780000")
    fLs.Cells(baris + 3, 3) = DtSPTPPh22
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTPPh22_pph = get_PPh_SPT_ssp(tahun, "780000", "22")
    fLs.Cells(baris + 3, 4) = DtSPTPPh22_pph
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    
    'UKP / 100000
    baris = 27
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiPPh22 = get_data_akuntansi_PPh22(tahun, "780000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiPPh22
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTPPh22 = get_data_SPTPPh22(tahun, "780000")
    fLs.Cells(baris + 3, 3) = DtSPTPPh22
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTPPh22_pph = get_PPh_SPT_ssp(tahun, "780000", "22")
    fLs.Cells(baris + 3, 4) = DtSPTPPh22_pph
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    


    fl.ActiveWorkbook.Save
    
    'fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub



Sub pph23_sheet_04(fileSimpan As String, DIVISI As String, tahun As String, _
            Masa_Pajak As String, kpp As String)
    Dim sql As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer, nmbulan
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long, bln As Integer
    Dim Total1 As Currency, totalPBKonstruksi As Currency, totalPBSewa As Currency
    Dim data1, data2, isi
    Dim DtAkuntansiPPh23 As Currency, DtSPTPPh23 As Currency, DtSPTPPh23_pph  As Currency
    
    f = fileSimpan
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Call setListInfo(Me.List1, "Proses Sheet04 - file not found")
            Exit Sub
        End If
    Else
        MsgBox "File tidak ditemukan", vbCritical
        Call setListInfo(Me.List1, "Proses Sheet04 - file not found")
            Exit Sub
        Exit Sub
    End If
    
    'open sheet 4, isi data looping
    Set fLs = fl.Sheets(4)
    baris = 3
    kolom = 2
    
    fLs.Cells(baris, kolom) = Me.cb_pph & " - Tahun " & tahun
    Call setListInfo(Me.List1, "Proses Sheet04 - Load Laporan")
    
    nmbulan = Array("", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", _
                    "Agustus", "September", "Oktober", "November", "Desember")
    
    'sewa
    'gedung1 / 400000 && gedung2 / 500000
    baris = 3
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiPPh23 = get_data_akuntansi_PPh23(tahun, "400000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiPPh23
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTPPh23 = get_data_SPTPPh23(tahun, "400000")
    fLs.Cells(baris + 3, 3) = DtSPTPPh23
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTPPh23_pph = get_PPh_SPT_ssp(tahun, "200000", "23") + get_PPh_SPT_ssp(tahun, "200000", "26")
    fLs.Cells(baris + 3, 4) = DtSPTPPh23_pph
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    DtAkuntansiPPh23 = get_data_akuntansi_PPh23(tahun, "500000")
    fLs.Cells(baris + 2, 8) = DtAkuntansiPPh23
    fLs.Cells(baris + 2, 8).NumberFormat = "#,##0"
    DtSPTPPh23 = get_data_SPTPPh21(tahun, "500000")
    fLs.Cells(baris + 3, 8) = DtSPTPPh23
    fLs.Cells(baris + 3, 8).NumberFormat = "#,##0"
    DtSPTPPh23_pph = get_PPh_SPT_ssp(tahun, "200000", "23") + get_PPh_SPT_ssp(tahun, "200000", "26")
    fLs.Cells(baris + 3, 9) = DtSPTPPh23_pph
    fLs.Cells(baris + 3, 9).NumberFormat = "#,##0"
    
    
    'infra1 / 200000 && infra2 / 300000
    baris = 11
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiPPh23 = get_data_akuntansi_PPh23(tahun, "200000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiPPh23
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTPPh23 = get_data_SPTPPh23(tahun, "200000")
    fLs.Cells(baris + 3, 3) = DtSPTPPh23
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTPPh23_pph = get_PPh_SPT_ssp(tahun, "200000", "23") + get_PPh_SPT_ssp(tahun, "200000", "26")
    fLs.Cells(baris + 3, 4) = DtSPTPPh23_pph
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    
    DtAkuntansiPPh23 = get_data_akuntansi_PPh23(tahun, "300000")
    fLs.Cells(baris + 2, 8) = DtAkuntansiPPh23
    fLs.Cells(baris + 2, 8).NumberFormat = "#,##0"
    DtSPTPPh23 = get_data_SPTPPh23(tahun, "300000")
    fLs.Cells(baris + 3, 8) = DtSPTPPh23
    fLs.Cells(baris + 3, 8).NumberFormat = "#,##0"
    DtSPTPPh23_pph = get_PPh_SPT_ssp(tahun, "200000", "23") + get_PPh_SPT_ssp(tahun, "200000", "26")
    fLs.Cells(baris + 3, 9) = DtSPTPPh23_pph
    fLs.Cells(baris + 3, 9).NumberFormat = "#,##0"
    
    'EPC / 780000
    baris = 19
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiPPh23 = get_data_akuntansi_PPh23(tahun, "780000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiPPh23
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTPPh23 = get_data_SPTPPh23(tahun, "780000")
    fLs.Cells(baris + 3, 3) = DtSPTPPh23
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTPPh23_pph = get_PPh_SPT_ssp(tahun, "200000", "23") + get_PPh_SPT_ssp(tahun, "200000", "26")
    fLs.Cells(baris + 3, 4) = DtSPTPPh23_pph
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
     
    
    'UKP / 100000
    baris = 27
    fLs.Cells(baris, 4) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    fLs.Cells(baris, 9) = "Per " & nmbulan(Masa_Pajak) & " " & tahun
    
    DtAkuntansiPPh23 = get_data_akuntansi_PPh23(tahun, "780000")
    fLs.Cells(baris + 2, 3) = DtAkuntansiPPh23
    fLs.Cells(baris + 2, 3).NumberFormat = "#,##0"
    DtSPTPPh23 = get_data_SPTPPh23(tahun, "780000")
    fLs.Cells(baris + 3, 3) = DtSPTPPh23
    fLs.Cells(baris + 3, 3).NumberFormat = "#,##0"
    DtSPTPPh23_pph = get_PPh_SPT_ssp(tahun, "200000", "23") + get_PPh_SPT_ssp(tahun, "200000", "26")
    fLs.Cells(baris + 3, 4) = DtSPTPPh23_pph
    fLs.Cells(baris + 3, 4).NumberFormat = "#,##0"
    


    fl.ActiveWorkbook.Save
    
    'fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
End Sub


Function get_nilaiAkunAll_divisi(noAkun As String, tahun As String, DIVISI As String) As Currency
    Dim sql As String, t As String
    
    If Trim(UCase(DIVISI)) = "ALL" Then
        DIVISI = ""
    End If
    sql = "select F_get_nilaiAkunAll_divisi('" & Trim(noAkun) & "'," & _
            Trim(tahun) & ",'" & Trim(DIVISI) & "')"
    t = cari_data1(cnn, sql, True)
    get_nilaiAkunAll_divisi = cek_Money(t)
End Function

Function get_nilaiAkun_divisi_bln(noAkun As String, tahun As String, bulan As String, _
            DIVISI As String) As Currency
    
    Dim sql As String, t As String
    
    If Trim(UCase(DIVISI)) = "ALL" Then
        DIVISI = ""
    End If
    sql = "select F_get_nilaiAkun_divisi_bln('" & Trim(noAkun) & "'," & _
            Trim(tahun) & ",'" & bulan & "','" & Trim(DIVISI) & "')"
    t = cari_data1(cnn, sql, True)
    get_nilaiAkun_divisi_bln = cek_Money(t)
End Function
    


Sub proses_data()
    Dim sql As String, t1 As String
    Dim fileSimpan As String, File1 As String
    Dim nmFile As String
    
    'On Error GoTo er1
    MsgBox "Pada saat proses selesai, kadang kala file excel tidak mau tampil. " & vbCr & _
            "Lakukan alt+tab. Jika ada konfirmasi, pilih 'switch'", vbInformation
    Me.disable_Form
    
    
    If Trim(Me.cb_pph.text) = "SPT 4(2) Sewa / Konstruksi" Then
        t1 = "spt4(2)"
    ElseIf Trim(Me.cb_pph.text) = "SPT 21" Then
        t1 = "spt21"
    ElseIf Trim(Me.cb_pph.text) = "SPT 22" Then
        t1 = "spt22"
    ElseIf Trim(Me.cb_pph.text) = "SPT 23" Then
        t1 = "spt23"
    Else
        t1 = "error"
    End If
    
    nmFile = "u" & Me.cb_divisi.text & "_" & Trim(Me.txt_Tahun) & Trim(Me.txt_Bulan) & _
            "kpp" & get_kode_combo(Me.cb_kpp, "#") & "_" & t1
    
    fileSimpan = App.Path & "\exp\" & Trim(nmFile) & ".xls"
    
    Me.List1.Clear
    Call setListInfo(Me.List1, "File simpan:" & fileSimpan)
    Call setListInfo(Me.List1, "Load Sheet1")
    
    
    If Trim(Me.cb_pph.text) = "SPT 4(2) Sewa / Konstruksi" Then
        'masukkan ke xls template
        Call setListInfo(Me.List1, "Proses Sheet01..")
        Call pph42_sheet_01(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses Sheet01..OK")
        Call delay(2)
            
        Call setListInfo(Me.List1, "Proses Sheet02..")
        Call pph42_sheet_02(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses Sheet02..OK")
        
        Call setListInfo(Me.List1, "Proses Sheet03..")
        Call pph42_sheet_03(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses Sheet03..OK")
        
        Call setListInfo(Me.List1, "Proses Sheet04..")
        Call pph42_sheet_04(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses Sheet04..OK")
    ElseIf Trim(Me.cb_pph.text) = "SPT 21" Then
        'masukkan ke xls template
        Call setListInfo(Me.List1, "Proses SPT 21 Sheet01..")
        Call pph21_sheet_01(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses SPT 21 Sheet01..OK")
        Call delay(2)
            
        Call setListInfo(Me.List1, "Proses SPT 21 Sheet02..")
        Call pph21_sheet_02(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses SPT 21 Sheet02..OK")
        
        Call setListInfo(Me.List1, "Proses SPT 21 Sheet03..")
        Call pph21_sheet_03(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses SPT 21 Sheet03..OK")
        
        Call setListInfo(Me.List1, "Proses SPT 21 Sheet04..")
        Call pph21_sheet_04(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses SPT 21 Sheet04..OK")
    ElseIf Trim(Me.cb_pph.text) = "SPT 22" Then
        'masukkan ke xls template
        Call setListInfo(Me.List1, "Proses SPT 22 Sheet01..")
        Call pph22_sheet_01(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses SPT 22 Sheet01..OK")
        Call delay(2)
            
        Call setListInfo(Me.List1, "Proses SPT 22 Sheet02..")
        Call pph22_sheet_02(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses SPT 22 Sheet02..OK")
        
        Call setListInfo(Me.List1, "Proses SPT 22 Sheet03..")
        Call pph22_sheet_03(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses SPT 22 Sheet03..OK")
        
        Call setListInfo(Me.List1, "Proses SPT 22 Sheet04..")
        Call pph22_sheet_04(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses SPT 22 Sheet04..OK")
    ElseIf Trim(Me.cb_pph.text) = "SPT 23" Then
        'masukkan ke xls template
        Call setListInfo(Me.List1, "Proses SPT 23 Sheet01..")
        Call pph23_sheet_01(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses SPT 23 Sheet01..OK")
        Call delay(2)
            
        Call setListInfo(Me.List1, "Proses SPT 23 Sheet02..")
        Call pph23_sheet_02(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses SPT 23 Sheet02..OK")
        
        Call setListInfo(Me.List1, "Proses SPT 23 Sheet03..")
        Call pph23_sheet_03(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses SPT 23 Sheet03..OK")
        
        Call setListInfo(Me.List1, "Proses SPT 23 Sheet04..")
        Call pph23_sheet_04(fileSimpan, Left(Trim(Me.cb_divisi), 6), Me.txt_Tahun, _
                    Me.txt_Bulan, get_kode_combo(Me.cb_kpp, "#"))
        Call setListInfo(Me.List1, "Proses SPT 23 Sheet04..OK")
    Else
    End If
    
    
    
    'open file
    'open by explorer
    File1 = "explorer.exe " & fileSimpan
    Call Shell(File1, vbNormalFocus)
    
    'done
    Me.Enable_Form
    MsgBox "Proses export selesai. " & vbCr & "File di :" & fileSimpan, vbInformation
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
End Sub


Private Sub cmd_load_Click()
    On Error GoTo er1
    Call proses_data
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
  nama_data = "Ekualisasi PPh"
  Call dbMySQL_open
    
  'load combo
  
  Me.txt_Tahun = CStr(Year(Now) - 1)
  Me.txt_Bulan = "12"
  
  Call load_Divisi(Me.cb_divisi, False, 1, True)
  Call load_KPP(Me.cb_kpp, False, 1, True)
  
  Me.cb_pph.Clear
  Me.cb_pph.AddItem "SPT 4(2) Sewa / Konstruksi"
  Me.cb_pph.AddItem "SPT 21"
  Me.cb_pph.AddItem "SPT 22"
  Me.cb_pph.AddItem "SPT 23"
  Me.cb_pph.ListIndex = 0
  
  
  Me.Height = 5010
  Me.Width = 12420
  
  'get level
  '2 : operator gedung
  '3 : UKP
  
  '---------
  
  Level1 = tbPengguna_getLevel1(frMenu1.nmLogin)
  If Level1 = 2 Then
    Me.cb_divisi.text = tbPengguna_getDivisi(frMenu1.nmLogin)
    Me.cb_divisi.Enabled = False
  ElseIf Level1 = 3 Then
  Else
    Call pesan2("Level tidak valid", , vbYellow)
   Me.cb_divisi.Enabled = False
  End If
 
  'Call LoadGrid
  'Call pesan2("Pilih Filter dan klik 'LOAD', atau " & vbCr & _
                "klik cari data dan ENTER")
End Sub


Private Sub Form_Resize()
    'If Me.Width - 405 > 0 Then Me.Frame3.Width = Me.Width - 405
    'If Me.Height - 2595 > 0 Then Me.Frame3.Height = Me.Height - 2595

    'If Me.Width - 645 > 0 Then Me.DataGrid1.Width = Me.Width - 645
    'If Me.Height - 3435 > 0 Then Me.DataGrid1.Height = Me.Height - 3435

    'If Me.Height - 3090 > 0 Then Me.txt_cari.Top = Me.Height - 3090
    'Me.Label6.Top = Me.txt_cari.Top
    'Me.cmd_export.Top = Me.txt_cari.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call dbMySQL_close
End Sub

Private Sub mncek_Click()
    Dim t As Currency
    Dim t2 As String
    
    t = get_nilaiAkunAll_divisi("51851", "2020", "300000")
    t2 = InputBox("", "", Format(t, "###,###"))
    t = get_nilaiAkunAll_divisi("83251", "2020", "300000")
    t2 = InputBox("", "", Format(t, "###,###"))
End Sub
