VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frm_BuktiPotong 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7245
   ClientLeft      =   225
   ClientTop       =   735
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_Stop 
      BackColor       =   &H008080FF&
      Caption         =   "Stop Load"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Height          =   4695
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   12015
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3855
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   6800
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton opt_print 
         Caption         =   "Print"
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   4260
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton opt_file 
         Caption         =   "File"
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   4260
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmd_BuktiPotong 
         Caption         =   "Report &Bukti Potong"
         Height          =   375
         Left            =   9960
         TabIndex        =   17
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Target Bukti Potong : "
         Height          =   210
         Left            =   240
         TabIndex        =   18
         Top             =   4282
         Visible         =   0   'False
         Width           =   1530
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " 2. Masa / KPP "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5640
      TabIndex        =   12
      Top             =   600
      Width           =   6495
      Begin VB.CommandButton cmd_Load 
         Caption         =   "3. &Load"
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cb_kpp 
         Height          =   330
         Left            =   840
         TabIndex        =   7
         Text            =   "x"
         Top             =   360
         Width           =   5415
      End
      Begin VB.TextBox txt_tahun 
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "KPP"
         Height          =   210
         Left            =   240
         TabIndex        =   14
         Top             =   420
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   795
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 1. Karyawan "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   5415
      Begin VB.TextBox txt_Nama 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txt_Nik 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txt_Npwp 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   390
         Width           =   3495
      End
      Begin VB.CheckBox ch_Nama 
         Caption         =   "Nama"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1050
         Width           =   735
      End
      Begin VB.CheckBox ch_Nik 
         Caption         =   "NIK"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   690
         Width           =   855
      End
      Begin VB.CheckBox ch_Npwp 
         Caption         =   "NPWP"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   6990
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10795
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10795
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CR 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lb_caption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1721-Bukti Potong PPh21 - 1721 A1 "
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
End
Attribute VB_Name = "frm_BuktiPotong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rS3 As ADODB.Recordset
Dim stopLoad As Boolean
Dim cnnTemp As ADODB.connection

Sub disable_Form()
    Me.Frame1.Enabled = False
    Me.Frame2.Enabled = False
    Me.Frame3.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame1.Enabled = True
    Me.Frame2.Enabled = True
    Me.Frame3.Enabled = True
End Sub


Function generate_sql() As String
    Dim sql As String, kondisi As String
    Dim cari As String
    
    'pph21tahunan2
        '0  No1, Bulan, Tahun, " & _
        '3  NPWP_KPP, kdPROYEK, kdCENTER, " & _
        '6  Nama, NPWP, NIK, " & _
        '9  Alamat, Jabatan, P_L, " & _
        '12 PTKP, Gaji, Tnj_PPh, " & _
        '15 Tunjangan_Lain, JHT_JPN, Bruto, " & _
        '18 Insentif, THR, Lainnya, " & _
        '21 Pensiun_Potongan_Lain, id1, tglupdate
        '24 biaya_jabatan, penghasilan_netto_sblmnya, pph21_terutang_sblmnya
    '----------
    
    sql = "select No1, Bulan, Tahun, " & _
            "NPWP_KPP, kdPROYEK, kdCENTER, " & _
            "Nama, NPWP, NIK, " & _
            "Alamat, Jabatan, P_L, " & _
            "PTKP, Gaji, Tnj_PPh, " & _
            "Tunjangan_Lain, JHT_JPN, Bruto, " & _
            "Insentif, THR, Lainnya, " & _
            "Pensiun_Potongan_Lain, id1, tglupdate, " & _
            "biaya_jabatan, penghasilan_netto_sblmnya, pph21_terutang_sblmnya from pph21tahunan2 "
    
    'cari
    cari = ""
    If Me.ch_Npwp.Value = "1" Then
        If Trim(cari) = "" Then
            cari = cari & " NPWP like '%" & Trim(Me.txt_Npwp) & "%' "
        Else
            cari = cari & " or NPWP like '%" & Trim(Me.txt_Npwp) & "%' "
        End If
    End If
    
    If Me.ch_Nik.Value = "1" Then
        If Trim(cari) = "" Then
            cari = cari & " NIK like '%" & Trim(Me.txt_Nik) & "%' "
        Else
            cari = cari & " or NIK like '%" & Trim(Me.txt_Nik) & "%' "
        End If
    End If
    
    If Me.ch_Nama.Value = "1" Then
        If Trim(cari) = "" Then
            cari = cari & " Nama like '%" & Trim(Me.txt_Nama) & "%' "
        Else
            cari = cari & " or Nama like '%" & Trim(Me.txt_Nama) & "%' "
        End If
    End If
    '---------
    
    '-- kondisi
    kondisi = ""
    If Trim(Me.cb_KPP.text) = "ALL" Or Trim(Me.cb_KPP.text) = "" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " NPWP_KPP = '" & get_kode_combo(Me.cb_KPP, "#") & "'"
    End If
    
    If Trim(Me.txt_tahun) = "" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Tahun = '" & Trim(Me.txt_tahun) & "'"
    End If
    '----
        
    If Trim(kondisi) <> "" Then
        sql = sql & " where " & kondisi
    Else
        sql = sql & " where true "
    End If
    
    If Trim(cari) <> "" Then
        sql = sql & " and (" & Trim(cari) & ")"
    End If
    
    sql = sql & " order by Nama, NIK, NPWP, Tahun, Bulan "
    'sql = InputBox("", "", sql)
    generate_sql = sql
    
End Function

Function generate_sql2() As String
    Dim sql As String, kondisi As String
    Dim cari As String
        
    sql = "select nomor as NoBuktiPotong, tahun, bulan_awal as awal, " & _
            "bulan_akhir as akhir, npwp_pemotong as NPWPKPP, nama_pemotong as KPP, " & _
            "npwp, NIK, Nama, " & _
            "Alamat, Jenis_kelamin as JK, ptkp, " & _
            "jabatan, no_1 as GajiPensiunTht, no_2 as TunjPPh, " & _
            "no_3 as TunjLain, no_4 as Honor, no_5 as Premi, " & _
            "no_6 as Lain, no_7 as BonusThr, no_8 as 1sd7, " & _
            "no_9 as jabPensiun, no_10 as IuranPensiun, no_11 as 9sd10, " & _
            "no_12 as net811, no_13 as netPajakSblm, no_14 as netHitungPPh, " & _
            "no_15 as PenghasilanTKP, no_16 as pkpSetahun, no_17 as pphSetahun, " & _
            "no_18 as PPhTelahDiptg, no_19 as pphTerutang, no_20 as PPhTlhDipotong, " & _
            "kdCENTER " & _
            "from buktipotong "
    
    'cari
    cari = ""
    If Me.ch_Npwp.Value = "1" Then
        If Trim(cari) = "" Then
            cari = cari & " NPWP like '%" & Trim(Me.txt_Npwp) & "%' "
        Else
            cari = cari & " or NPWP like '%" & Trim(Me.txt_Npwp) & "%' "
        End If
    End If
    
    If Me.ch_Nik.Value = "1" Then
        If Trim(cari) = "" Then
            cari = cari & " NIK like '%" & Trim(Me.txt_Nik) & "%' "
        Else
            cari = cari & " or NIK like '%" & Trim(Me.txt_Nik) & "%' "
        End If
    End If
    
    If Me.ch_Nama.Value = "1" Then
        If Trim(cari) = "" Then
            cari = cari & " Nama like '%" & Trim(Me.txt_Nama) & "%' "
        Else
            cari = cari & " or Nama like '%" & Trim(Me.txt_Nama) & "%' "
        End If
    End If
    '---------
    
    '-- kondisi
    kondisi = ""
    If Trim(Me.cb_KPP.text) = "ALL" Or Trim(Me.cb_KPP.text) = "" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " npwp_pemotong = '" & get_kode_combo(Me.cb_KPP, "#") & "'"
    End If
    
    If Trim(Me.txt_tahun) = "" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Tahun = '" & Trim(Me.txt_tahun) & "'"
    End If
    '----
        
    If Trim(kondisi) <> "" Then
        sql = sql & " where " & kondisi
    Else
        sql = sql & " where true "
    End If
    
    If Trim(cari) <> "" Then
        sql = sql & " and (" & Trim(cari) & ")"
    End If
    
    sql = sql & " order by Nama, NIK, NPWP, Tahun, Bulan_awal "
    'sql = InputBox("", "", sql)
    generate_sql2 = sql
    
End Function

Sub format_Grid()
    
    Dim jenisPPh As String
    Dim jRec As Long
    Dim c As Integer
    
    jRec = RecordCount(rS3)
    If jRec <= 0 Then Exit Sub
    
    
    'pph21tahunan2
        '0  No1, Bulan, Tahun, " & _
        '3  NPWP_KPP, kdPROYEK, kdCENTER, " & _
        '6  Nama, NPWP, NIK, " & _
        '9  Alamat, Jabatan, P_L, " & _
        '12 PTKP, Gaji, Tnj_PPh, " & _
        '15 Tunjangan_Lain, JHT_JPN, Bruto, " & _
        '18 Insentif, THR, Lainnya, " & _
        '21 Pensiun_Potongan_Lain, id1, tglupdate
        '24 biaya_jabatan, penghasilan_netto_sblmnya, pph21_terutang_sblmnya
    '----------
        
        For c = 0 To rS3.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
            
            If c = 0 Or c = 1 Or c = 2 Or c = 4 Or c = 5 Or c = 11 Or c = 12 Or c = 22 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).Width = 700
            End If
            
            
            If c = 23 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 13 Or c = 14 Or c = 15 Or c = 16 Or c = 17 Or c = 18 Or c = 19 Or _
                c = 20 Or c = 21 Or c = 24 Or c = 25 Or c = 26 Then
                
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next

End Sub

Private Sub cb_kpp_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 113 And Shift = 0 Then Call load_KPP(Me.cb_KPP, True)
End Sub

Sub export_data(fl As Object, npwp As String, NIK As String, nama As String, npwpKpp As String, _
                tahun As String, id1 As String)
                
    Dim t As String, ptkp As String, t2 As String
    Dim penghasilanKenaPajakSetahun As Currency
    Dim cq61 As Currency, cq62 As Currency
    Dim npwp_kpp As String
    
    'id1 = id data
    fl.Cells(11, "BF") = "12"
    fl.Cells(11, "BN") = Right(tahun, 2)
    fl.Cells(11, "DB") = "'" & adddigit(CLng(tbPph21Tahunan2_getBulanAwal(npwp, NIK, nama, tahun, npwpKpp)), 2)
    fl.Cells(11, "DJ") = "'" & adddigit(tbPph21Tahunan2_getBulanAkhir(npwp, NIK, nama, tahun, npwpKpp), 2)
    
    npwp_kpp = tbPph21Tahunan2_getData_byId(id1, "NPWP_KPP")
    fl.Cells(16, "W") = "'" & format_Npwp_awal(Left(npwp_kpp, 9))
    fl.Cells(16, "BB") = "'" & Mid(npwp_kpp, 10, 3)
    fl.Cells(16, "BM") = "'" & Mid(npwp_kpp, 13, 3)
    
    fl.Cells(18, "W") = UCase(tbMKpp_get("nama", npwp_kpp))
    
    
    fl.Cells(24, "S") = "'" & format_Npwp_awal(Left(npwp, 9))
    fl.Cells(24, "AW") = "'" & Mid(npwp, 10, 3)
    fl.Cells(24, "BF") = "'" & Mid(npwp, 13, 3)
    
    fl.Cells(26, "S") = "'" & NIK
    fl.Cells(29, "S") = UCase(nama)
    fl.Cells(32, "S") = Trim(Left(UCase(tbPph21Tahunan2_getData_byId(id1, "Alamat")), 50))
    
    If tbPph21Tahunan2_getData_byId(id1, "P_L") = "P" Then
        fl.Cells(36, "Y") = ""
        fl.Cells(36, "AS") = "X"
    Else
        fl.Cells(36, "Y") = "X"
        fl.Cells(36, "AS") = ""
    End If
    
    ptkp = tbPph21Tahunan2_getData_byId(id1, "PTKP")
    t = Replace(ptkp, "/", "")
    t = Replace(t, " ", "")
    t = Replace(t, "0", "")
    t2 = Right(t, 1)
    If IsNumeric(t2) = False Then t2 = "0"
    If Trim(t2) = "" Then t2 = "0"
    If Left(t, 1) = "K" Then
        fl.Cells(26, "BT") = t2
        fl.Cells(26, "CK") = ""
        fl.Cells(26, "DA") = ""
    ElseIf Left(t, 1) = "T" Then
        fl.Cells(26, "BT") = ""
        fl.Cells(26, "CK") = t2
        fl.Cells(26, "DA") = ""
    Else
        fl.Cells(26, "BT") = ""
        fl.Cells(26, "CK") = ""
        fl.Cells(26, "DA") = t2
    End If
    
    fl.Cells(29, "CK") = tbPph21Tahunan2_getJabatan(npwp, NIK, nama, tahun)
    
    'totalbruto = sum(Gaji + Tnj_PPh + JHT_JPN + Tunjangan_Lain + Insentif)
    fl.Cells(46, "CQ") = tbPph21Tahunan2_getTotal("Gaji", npwp, NIK, nama, tahun, npwpKpp)
    fl.Cells(47, "CQ") = tbPph21Tahunan2_getTotal("Tnj_PPh", npwp, NIK, nama, tahun, npwpKpp)
    fl.Cells(48, "CQ") = tbPph21Tahunan2_getTotal("Tunjangan_Lain", npwp, NIK, nama, tahun, npwpKpp)
    fl.Cells(50, "CQ") = tbPph21Tahunan2_getTotal("Tunjangan_Lain + JHT_JPN", npwp, NIK, nama, tahun, npwpKpp)
    fl.Cells(52, "CQ") = tbPph21Tahunan2_getTotal("Insentif + THR + Lainnya", npwp, NIK, nama, tahun, npwpKpp)
            
    fl.Cells(55, "CQ") = tbPph21Tahunan2_getTotalBiayaJabatan(npwp, NIK, nama, tahun, npwpKpp)

    
    fl.Cells(56, "CQ") = tbPph21Tahunan2_getTotal("Pensiun_Potongan_Lain", npwp, NIK, nama, tahun, npwpKpp)
    cq61 = cek_Money(fl.Cells(61, "CQ"))
    fl.Cells(62, "CQ") = tbM_Ptkp_getNilai(ptkp)
    cq62 = cek_Money(fl.Cells(62, "CQ"))
    
    If cq61 - cq62 > 0 Then
        fl.Cells(63, "CQ") = NearestThousand(cq61 - cq62)
    Else
        fl.Cells(63, "CQ") = "0"
    End If
    
    penghasilanKenaPajakSetahun = cek_Money(fl.Cells(63, "CQ"))
    
    fl.Cells(64, "CQ") = get_pph21Setahun(penghasilanKenaPajakSetahun)
    fl.Cells(66, "CQ") = fl.Cells(64, "CQ")
    fl.Cells(67, "CQ") = fl.Cells(64, "CQ")
    fl.Cells(72, "Q") = "'09.321.683.6"
    fl.Cells(72, "AW") = "'411"
    
    fl.Cells(75, "Q") = "FARID FACHRUR RAZI"
    fl.Cells(75, "CL") = tahun
    
End Sub



Private Sub cmd_BuktiPotong_Click()
    Dim p
    Dim c As Long, jRec As Long, mod1 As Long
    Dim sql As String
    Dim adaError As Boolean
    
    stopLoad = False
    Me.cmd_Stop.Visible = True
       
    jRec = RecordCount(rS3)
    If jRec <= 0 Then
        MsgBox "tidak ada data", vbCritical
        Exit Sub
    End If
    Call create_ds_Access("c:\dbpph.dsn", App.Path & "\data\", App.Path & "\data\dbrep.mdb")
    
    'open db_rep
    'delete temporary
    Call db_Access_open(cnnTemp, App.Path & "\data\dbrep.mdb")
    
    'delete
    sql = "delete from buktipotong"
    If ExecSQL1(cnnTemp, sql) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    '-----------
        
    p = MsgBox("Hitung ulang ?", vbYesNo)
    If p = vbYes Then
        rS3.MoveFirst
        c = 1
        Do While rS3.EOF = False
            If stopLoad = True Then Exit Do
            Call info(1, "Load data. Run " & c & "/" & jRec & " -- " & Round((c / jRec) * 100, 2) & _
                        "%", Me.StatusBar1)
            mod1 = c Mod 2000
            If mod1 = 0 Then Call dbMySQL_open
            
            Call fetch_Bukti_Potong(Me.txt_tahun.text, rS3, cnnTemp, Me.StatusBar1, adaError)
            If adaError = True Then Exit Do
            rS3.MoveNext
            c = c + 1
        Loop
    Else
        'copy dari tabel buktipotong
        Call fetch_Bukti_Potong_langsung
    End If
        
    Call tampil_report(CR, App.Path & "\rep\rep_Bukti_Potong.rpt", 85)
    Call db_access_close(cnnTemp)
    
End Sub

Sub fetch_Bukti_Potong_langsung()
    Dim sql As String
    Dim rS2 As ADODB.Recordset
    Dim jRec As Long, c As Long, mod1 As Long
    
    Dim npwp_kpp As String, adaError As Boolean
    
    Dim alamat As String, bulan_akhir As String, bulan_awal As String, Jabatan As String
    Dim jenis_kelamin As String, nama As String, Nama_Pemotong As String, nama_ttd As String
    Dim NIK As String, no_1 As Currency, no_10 As Currency, no_11 As Currency, no_12 As Currency
    Dim no_13 As Currency, no_14 As Currency, no_15 As Currency, no_16 As Currency, no_17 As Currency
    Dim no_18 As Currency, no_19 As Currency, no_2 As Currency, no_20 As Currency, no_3 As Currency
    Dim no_4 As Currency, no_5 As Currency, no_6 As Currency, no_7 As Currency, no_8 As Currency
    Dim no_9 As Currency, nomor As String, npwp As String
    Dim npwp_ttd As String, ptkp As String, tahun As String, NPWP_Pemotong As String
    
    Dim no_13s As Currency, no_19s As Currency, jmlBulan As Integer
    Dim kdCENTER As String
    
    sql = generate_sql2
    If OpenRecordSet(cnn, rS2, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rS2)
    If jRec <= 0 Then
        MsgBox "tidak ada data bukti potong. Harus hitung ulang", vbInformation
        Exit Sub
    End If
        
    rS2.MoveFirst
    Do While rS2.EOF = False
        Call info(1, "run move data. " & c & "/" & jRec & " -- " & Round((c / jRec) * 100, 2) & "%" _
                    , Me.StatusBar1)
        mod1 = c Mod 4000
        If mod1 = 0 Then Call dbMySQL_open
        
        '-- referensi
    '0: nomor as NoBuktiPotong, tahun, bulan_awal as awal, " & _
    '3: "bulan_akhir as akhir, npwp_pemotong as NPWPKPP, nama_pemotong as KPP, " & _
    '6: "npwp, NIK, Nama, "
    
        nomor = cek_null(rS2(0))
        tahun = cek_null(rS2(1))
        bulan_awal = cek_null(rS2(2))
        bulan_akhir = cek_null(rS2(3))
        NPWP_Pemotong = cek_null(rS2(4))
        npwp_kpp = NPWP_Pemotong
        Nama_Pemotong = cek_null(rS2(5))
        npwp = cek_null(rS2(6))
        
        '9: "Alamat, Jenis_kelamin as JK, ptkp, " & _
    '12: "jabatan, no_1 as GajiPensiunTht, no_2 as TunjPPh, " & _
    '15: "no_3 as TunjLain, no_4 as Honor, no_5 as Premi, " & _

        NIK = cek_null(rS2(7))
        nama = cek_null(rS2(8))
        alamat = cek_null(rS2(9))
        jenis_kelamin = cek_null(rS2(10))
        ptkp = cek_null(rS2(11))
        Jabatan = cek_null(rS2(12))
        no_1 = cek_null(rS2(13))
        no_2 = cek_null(rS2(14))
        no_3 = cek_null(rS2(15))
        no_4 = cek_null(rS2(16))
        no_5 = cek_null(rS2(17))
        
        '18: "no_6 as Lain, no_7 as BonusThr, no_8 as 1sd7, " & _
    '21: "no_9 as jabPensiun, no_10 as IuranPensiun, no_11 as 9sd10, " & _
    '24: "no_12 as net811, no_13 as netPajakSblm, no_14 as netHitungPPh, " & _

        no_6 = cek_null(rS2(18))
        no_7 = cek_null(rS2(19))
        no_8 = cek_null(rS2(20))
        no_9 = cek_null(rS2(21))
        no_10 = cek_null(rS2(22))
        no_11 = cek_null(rS2(23))
        no_12 = cek_null(rS2(24))
        
        '27: "no_15 as PenghasilanTKP, no_16 as pkpSetahun, no_17 as pphSetahun, " & _
    '30: "no_18 as PPhTelahDiptg, no_19 as pphTerutang, no_20 as PPhTlhDipotong, " & _
    '33: "kdCENTER " & _
    '----
        
        no_13 = cek_null(rS2(25))
        no_14 = cek_null(rS2(26))
        no_15 = cek_null(rS2(27))
        no_16 = cek_null(rS2(28))
        no_17 = cek_null(rS2(29))
        no_18 = cek_null(rS2(30))
        no_19 = cek_null(rS2(31))
        no_20 = cek_null(rS2(32))
        kdCENTER = cek_null(rS2(33))
        npwp_ttd = "09.321.683.6411000"
        nama_ttd = "FARID FACHRUR RAZI"

        '--insert ke dbtemp
        '- jika di panggil melalui form rekap, insert ke db utama
        sql = "insert into buktipotong (nomor, tahun, bulan_awal, " & _
                "bulan_akhir, npwp_pemotong, nama_pemotong, " & _
                "npwp, NIK, Nama, " & _
                "Alamat, Jenis_kelamin, ptkp, " & _
                "jabatan, no_1, no_2, " & _
                "no_3, no_4, no_5, " & _
                "no_6, no_7, no_8, " & _
                "no_9, no_10, no_11, " & _
                "no_12, no_13, no_14, " & _
                "no_15, no_16, no_17, " & _
                "no_18, no_19, no_20, " & _
                "npwp_ttd, nama_ttd, kdCENTER) values ('" & _
                Trim(nomor) & "','" & Trim(tahun) & "','" & Trim(bulan_awal) & "','" & _
                Trim(bulan_akhir) & "','" & Trim(npwp_kpp) & "','" & Trim(Nama_Pemotong) & "','" & _
                Trim(npwp) & "','" & Trim(NIK) & "','" & Trim(nama) & "','" & _
                Trim(alamat) & "','" & Trim(jenis_kelamin) & "','" & Trim(ptkp) & "','" & _
                Trim(Jabatan) & "','" & Trim(no_1) & "','" & Trim(no_2) & "','" & _
                Trim(no_3) & "','" & Trim(no_4) & "','" & Trim(no_5) & "','" & _
                Trim(no_6) & "','" & Trim(no_7) & "','" & Trim(no_8) & "','" & _
                Trim(no_9) & "','" & Trim(no_10) & "','" & Trim(no_11) & "','" & _
                Trim(no_12) & "','" & Trim(no_13) & "','" & Trim(no_14) & "','" & _
                Trim(no_15) & "','" & Trim(no_16) & "','" & Trim(no_17) & "','" & _
                Trim(no_18) & "','" & Trim(no_19) & "','" & Trim(no_20) & "','" & _
                Trim(npwp_ttd) & "','" & Trim(nama_ttd) & "', '" & kdCENTER & "')"
        If ExecSQL1(cnnTemp, sql) <> 0 Then
            sql = InputBox("sql error", "", sql)
            adaError = True
            Exit Do
        End If
        '---
        
        rS2.MoveNext
        c = c + 1
    Loop
End Sub

Sub bukti_Potong_Xls()
    Dim fl As Object
    Dim id1 As String
    Dim rsKPP As ADODB.Recordset
    Dim sql As String, npwp As String, NIK As String, nama As String, npwp_kpp As String
    Dim jRec As Long, c As Long
    
    'cek apa file template ada
    'isi file template
    'copykan ke folder export
    
    On Error GoTo er1
    
    If RecordCount(rS3) <= 0 Then
        MsgBox "tidak ada data", vbCritical
        Exit Sub
    End If
    
    id1 = cek_null(rS3(22))
    
    '-- dari id yang ada, loop untuk semua NPWP_KPP
    
    '-----
    'pph21tahunan2
        '0  No1, Bulan, Tahun, " & _
        '3  NPWP_KPP, kdPROYEK, kdCENTER, " & _
        '6  Nama, NPWP, NIK, " & _
        '9  Alamat, Jabatan, P_L, " & _
        '12 PTKP, Gaji, Tnj_PPh, " & _
        '15 Tunjangan_Lain, JHT_JPN, Bruto, " & _
        '18 Insentif, THR, Lainnya, " & _
        '21 Pensiun_Potongan_Lain, id1, tglupdate
        '24 tunjangan_jab
    '----------
    npwp = cek_null(rS3(7))
    nama = cek_null(rS3(6))
    NIK = cek_null(rS3(8))
    
    sql = "select distinct NPWP_KPP from pph21tahunan2 where NPWP = '" & Trim(npwp) & _
            "' and NIK = '" & Trim(NIK) & "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(Me.txt_tahun) & "'"
    If OpenRecordSet(cnn, rsKPP, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error sql", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rsKPP)
    If jRec <= 0 Then
        MsgBox "tidak ada data KPP", vbInformation
        Exit Sub
    Else
        Call pesan2("data di " & jRec & " KPP", 1)
    End If
    
    rsKPP.MoveFirst
    c = 1
    Do While rsKPP.EOF = False
        Call info(1, "Load data. Run " & c & "/" & jRec, Me.StatusBar1)
        npwp_kpp = cek_null(rsKPP(0))
        Shell "cmd /c del " & App.Path & "\data\temp99.xls"
        Call CopyFileWindowsWay(App.Path & "\data\template bukti_potong.xls", App.Path & "\data\temp99.xls")
    
        '-- open xls dan update data
        If open_xls_lateBinding(fl, App.Path & "\data\temp99.xls") <> 0 Then
            MsgBox "Error open File template", vbCritical
        Else
            Call export_data(fl, npwp, NIK, nama, npwp_kpp, Trim(Me.txt_tahun), id1)
            fl.ActiveWorkbook.CheckCompatibility = False
            
            If Me.opt_file.Value = True Then
                fl.ActiveWorkbook.SaveAs App.Path & "\exp\bpot_" & _
                                           get_nama_file(npwp, NIK, nama, Trim(Me.txt_tahun), npwp_kpp) & _
                                           ".xls", FileFormat:=56
                MsgBox "Proses Export Selesai." & vbCr & _
                        "File ada di " & App.Path & "\exp\bpot_" & get_nama_file(npwp, NIK, nama, _
                                            Trim(Me.txt_tahun), npwp_kpp) & ".xls"
            Else
                fl.ActiveWorkbook.PrintOut
                fl.ActiveWorkbook.Save
            End If
        End If
        Call close_xls_lateBinding(fl)
        
    
        rsKPP.MoveNext
    Loop
    
    
    
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
    Call close_xls_lateBinding(fl)

End Sub


Function get_nama_file(npwp As String, NIK As String, nama As String, tahun As String, npwp_kpp As String) As String
    'npwpkpp_npwp_nama
    
    Dim nmFile As String
    
    
    nmFile = Right(tahun, 2) & "_" & Trim(Right(npwp_kpp, 6)) & "_" & Trim(Right(npwp, 6)) & "_" & Trim(Left(nama, 10))
    get_nama_file = nmFile
End Function

Private Sub cmd_load_Click()
    Dim sql As String
    Dim jRec As Long, c As Long, a As Integer
    Dim rs0 As ADODB.Recordset
    Dim npwp As String, NIK As String, nama As String, tahun As String, masa As String
    
    If Trim(Me.txt_tahun) = "" Then
        MsgBox "Tahun harus diisi", vbCritical
        Exit Sub
    End If
    
    Me.disable_Form
    Me.cmd_Stop.Visible = True
    stopLoad = False
    
    sql = generate_sql
    DoEvents
    'MsgBox sql
    If Trim(sql) <> "" Then
        Call dbMySQL_open
        If OpenRecordSet(cnn, rs0, sql, adOpenStatic, adLockPessimistic, adUseClient) <> 0 Then
            MsgBox "error open sql"
            Me.Enable_Form
            Exit Sub
        End If
        
        '-- manipulasi data
        If createRS_duplicate(rs0, rS3) = True Then
            jRec = RecordCount(rs0)
            If jRec > 0 Then
                'copykan isi rs ke rsGrid
                rs0.MoveFirst
                c = 1
                Do While rs0.EOF = False
                    Call info(1, "Load Grid | Copy | Run " & c & "/" & jRec & " -- " & _
                                Round((c / jRec) * 100, 2) & "%", Me.StatusBar1)
                    If stopLoad = True Then Exit Do
                    rS3.AddNew
                    For a = 0 To rs0.Fields.Count - 1
                        rS3.Fields(a) = cek_null(rs0(a))
                    Next
                    rS3.Update
                    c = c + 1
                    rs0.MoveNext
                Loop
                
            End If
            Set Me.DataGrid1.DataSource = rS3
            Me.Frame3.Caption = " Data | JmlData: " & _
                                RecordCount(rS3)
            Call info(1, "JumlahData: " & RecordCount(rS3), Me.StatusBar1)
                                
        Else
            Set Me.DataGrid1.DataSource = Nothing
            Me.Frame3.Caption = "  data | ERROR "
        End If
        
        
        Call format_Grid
    End If
    Me.Enable_Form
    Me.cmd_Stop.Visible = False
End Sub

Private Sub cmd_Stop_Click()
    stopLoad = True
    Me.cmd_Stop.Visible = False
End Sub


Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  Dim p
  
    
  Me.txt_tahun.text = Year(Now) - 1
  Me.txt_Nama = ""
  Me.txt_Nik = ""
  Me.txt_Npwp = ""
  Me.cmd_Stop.Visible = False
  
  '---------------------
    If dbMySQL_open = True Then
    Else
        MsgBox "Open Database Tidak Sukses", vbExclamation
    End If
    '---------------------
    
  'load combo
  Call load_KPP(Me.cb_KPP, False, 1)
  
  'get level
  '2 : operator gedung
  '3 : UKP
  
  '---------
  
  Level1 = tbPengguna_getLevel1(frMenu1.nmLogin)
  Call info(2, "Level " & Level1, Me.StatusBar1)
  
  
  
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call dbMySQL_close
End Sub

Private Sub opt_print_Click()
    If Me.opt_print.Value = True Then
        MsgBox "pastikan default printer telah di setting", vbInformation
    End If
End Sub

Private Sub txt_Nama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmd_load_Click
End Sub

Private Sub txt_Nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmd_load_Click
End Sub

Private Sub txt_Npwp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmd_load_Click
End Sub
