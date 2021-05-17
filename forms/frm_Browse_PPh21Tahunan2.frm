VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Browse_PPh21Tahunan2 
   ClientHeight    =   7245
   ClientLeft      =   240
   ClientTop       =   750
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
   ScaleHeight     =   7245
   ScaleWidth      =   12300
   Begin VB.Frame Frame3 
      Height          =   4815
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   12015
      Begin VB.CommandButton cmd_edit 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Edit Data"
         Height          =   375
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmd_hapus1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Hapus Data(s)"
         Height          =   375
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CommandButton cmd_Hapus 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Hapus Data 1 KPP"
         Height          =   375
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CommandButton cmd_export 
         Caption         =   "&Export XLS"
         Height          =   375
         Left            =   10920
         TabIndex        =   19
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txt_cari 
         Height          =   375
         Left            =   1200
         TabIndex        =   17
         Text            =   "Text1"
         ToolTipText     =   "input dan ENTER"
         Top             =   4320
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3975
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   7011
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cari Data "
         Height          =   210
         Left            =   240
         TabIndex        =   18
         Top             =   4402
         Width           =   705
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
      Height          =   1335
      Left            =   4200
      TabIndex        =   10
      Top             =   600
      Width           =   7935
      Begin VB.CommandButton cmd_Load 
         Caption         =   "3. &Load"
         Height          =   375
         Left            =   6600
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cb_kpp 
         Height          =   330
         Left            =   3000
         TabIndex        =   5
         Text            =   "x"
         Top             =   360
         Width           =   4695
      End
      Begin VB.TextBox txt_masa 
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txt_tahun 
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "KPP"
         Height          =   210
         Left            =   2520
         TabIndex        =   13
         Top             =   420
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Masa"
         Height          =   210
         Left            =   360
         TabIndex        =   12
         Top             =   915
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   360
         TabIndex        =   11
         Top             =   442
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 1. Divisi / Jenis PPh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   3975
      Begin VB.ComboBox cb_proyek 
         Height          =   330
         Left            =   1200
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox cb_divisi 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Text            =   "x"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Proyek"
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Divisi"
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   420
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   6990
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
      Caption         =   "Browse PPh 21 Tahunan(2)"
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
Attribute VB_Name = "frm_Browse_PPh21Tahunan2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset

Sub disable_Form()
    Me.Frame1.Enabled = False
    Me.Frame2.Enabled = False
    Me.Frame3.Enabled = False
    Me.cmd_Hapus.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame1.Enabled = True
    Me.Frame2.Enabled = True
    Me.Frame3.Enabled = True
    Me.cmd_Hapus.Enabled = True
End Sub


Function generate_sql() As String
    Dim sql As String, kondisi As String
    Dim jenisPPh As String, cari As String
    
    'kondisi
    kondisi = ""
    cari = ""
    
    If Trim(Me.cb_divisi.Text) = "ALL" Or Trim(Me.cb_divisi.Text) = "" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " kdCENTER = '" & get_kode_combo(Me.cb_divisi, "-") & "'"
    End If
    
    'proyek
    If Trim(Me.cb_proyek.Text) = "ALL" Or Trim(Me.cb_proyek.Text) = "" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " kdPROYEK = '" & get_kode_combo(Me.cb_proyek, "-") & "'"
    End If
    
    If Trim(Me.txt_tahun) = "" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Tahun = '" & Trim(Me.txt_tahun) & "'"
    End If
    
    If Trim(Me.txt_masa.Text) = "" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Bulan = '" & Trim(Me.txt_masa.Text) & "'"
    End If
    
    
    If Trim(Me.cb_KPP.Text) = "ALL" Or Trim(Me.cb_KPP.Text) = "" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " NPWP_KPP = '" & get_kode_combo(Me.cb_KPP, "#") & "'"
    End If
    
    
        sql = "select No1, Bulan, Tahun, " & _
            "NPWP_KPP, kdPROYEK, kdCENTER, " & _
            "Nama, NPWP, NIK, " & _
            "Alamat, Jabatan, P_L, " & _
            "PTKP, Gaji, Tnj_PPh, " & _
            "Tunjangan_Lain, JHT_JPN, Bruto, " & _
            "Insentif, THR, Lainnya, " & _
            "Pensiun_Potongan_Lain, penghasilan_netto_sblmnya, pph21_terutang_sblmnya, " & _
            "nrp, id1, tglupdate, " & _
            "biaya_jabatan, nilai_beban " & _
            "from pph21tahunan2"
        If Trim(Me.txt_cari.Text) <> "" Then
            cari = "Nama like '%" & Trim(Me.txt_cari.Text) & "%' or " & _
                    "NPWP like '%" & Trim(Me.txt_cari.Text) & "%' or " & _
                    "NIK like '%" & Trim(Me.txt_cari.Text) & "%' or " & _
                    "Alamat like '%" & Trim(Me.txt_cari.Text) & "%' "
        End If
    
    '-- fiter cari
    If Trim(kondisi) <> "" Then
        sql = sql & " where (" & kondisi & ") "
    End If
    
    If Trim(cari) <> "" Then
        If Trim(kondisi) = "" Then
            sql = sql & " where " & cari
        Else
            sql = sql & " and (" & cari & ") "
        End If
    End If
    
    generate_sql = sql & " order by cast(no1 as int), NPWP, NIK, Nama, Tahun, Bulan "
    'sql = InputBox("", "", sql)
End Function

Sub format_Grid()
    
    Dim jenisPPh As String
    Dim jRec As Long
    Dim c As Integer
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
    
    '-----
    'pph21tahunan2
        '0  No1, Bulan, Tahun, " & _
        '3  NPWP_KPP, kdPROYEK, kdCENTER, " & _
        '6  Nama, NPWP, NIK, " & _
        '9  Alamat, Jabatan, P_L, " & _
        '12 PTKP, Gaji, Tnj_PPh, " & _
        '15 Tunjangan_Lain, JHT_JPN, Bruto, " & _
        '18 Insentif, THR, Lainnya, " & _
        '21 Pensiun_Potongan_Lain, penghasilan_netto_sblmnya, pph21_terutang_sblmnya, " & _
        '24 nrp, id1, tglupdate, " & _
        '27 biaya_jabatan, nilai_beban " & _
    '----------
    
        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
            
            'kecil
            If c = 0 Or c = 1 Or c = 2 Or c = 4 Or c = 5 Or c = 11 Or c = 12 Or c = 25 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).Width = 700
            End If
    
            If c = 26 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 13 Or c = 14 Or c = 15 Or c = 16 Or c = 17 Or c = 18 Or c = 19 Or _
                c = 20 Or c = 21 Or c = 22 Or c = 23 Or c = 27 Or c = 28 Then
                
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next

End Sub

Private Sub cb_kpp_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 113 And Shift = 0 Then Call load_KPP(Me.cb_KPP, True)
End Sub

Private Sub cmd_edit_Click()
    '-----
    'pph21tahunan2
        '0  No1, Bulan, Tahun, " & _
        '3  NPWP_KPP, kdPROYEK, kdCENTER, " & _
        '6  Nama, NPWP, NIK, " & _
        '9  Alamat, Jabatan, P_L, " & _
        '12 PTKP, Gaji, Tnj_PPh, " & _
        '15 Tunjangan_Lain, JHT_JPN, Bruto, " & _
        '18 Insentif, THR, Lainnya, " & _
        '21 Pensiun_Potongan_Lain, penghasilan_netto_sblmnya, pph21_terutang_sblmnya, " & _
        '24 nrp, id1, tglupdate, " & _
        '27 biaya_jabatan, nilai_beban " & _
    '----------
    
    Dim nama As String, NIK As String, id1 As String, npwp As String
    Dim alamat As String, Jabatan As String, P_L As String, Bulan As String
    Dim p
    Dim penghasilan_netto_sblmnya As Currency, pph21_terutang_sblmnya As Currency
    
    Bulan = cek_null(rs(1))
    nama = cek_null(rs(6))
    
    NIK = cek_null(rs(8))
    id1 = cek_null(rs(25))
    npwp = cek_null(rs(7))
    alamat = cek_null(rs(9))
    Jabatan = cek_null(rs(10))
    P_L = cek_null(rs(11))
    penghasilan_netto_sblmnya = cek_Money(rs(22))
    pph21_terutang_sblmnya = cek_Money(rs(23))
    
    If Trim(id1) = "" Then Exit Sub
    p = MsgBox("Ubah data " & nama & " / " & NIK & "?", vbYesNo)
    If p = vbYes Then
        nama = InputBox("Input Nama", "", nama)
        npwp = InputBox("Input NPWP", "", npwp)
        NIK = InputBox("Input NIK", "", NIK)
        alamat = InputBox("Input Alamat", "", alamat)
        Jabatan = InputBox("Input Jabatan", "", Jabatan)
        P_L = InputBox("Input Jenis Kelamin", "", P_L)
        Bulan = InputBox("Input bulan", "", Bulan)
        penghasilan_netto_sblmnya = cek_Money(InputBox("penghasilan_netto_sblmnya", "", penghasilan_netto_sblmnya))
        pph21_terutang_sblmnya = cek_Money(InputBox("pph21_terutang_sblmnya", "", pph21_terutang_sblmnya))
        
        If Trim(nama) = "" Or Trim(NIK) = "" Or Trim(npwp) = "" Or Trim(alamat) = "" Or Trim(Jabatan) = "" _
            Or Trim(P_L) = "" Or Bulan = "" Then Exit Sub
    
        If tbPph21Tahunan2_Edit(id1, nama, npwp, NIK, alamat, Jabatan, P_L, Bulan, penghasilan_netto_sblmnya, _
                                pph21_terutang_sblmnya) = True Then
            Call cmd_Load_Click
        Else
            Call pesan2("error edit", 5, vbYellow)
        End If
    Else
        Call pesan2("Batal", 5, vbYellow)
    End If

End Sub

Private Sub cmd_export_Click()
    Dim jRec As Long
    Dim judul As String
    
    judul = ""
    
    If Trim(Me.cb_divisi) = "" Or Trim(Me.cb_divisi) = "ALL" Then
    Else
        judul = judul & "Divisi " & Me.cb_divisi & ", "
    End If
    
    If Trim(Me.cb_proyek) = "" Or Trim(Me.cb_proyek) = "ALL" Then
    Else
        judul = judul & "Proyek " & Me.cb_proyek & ", "
    End If
    
    If Trim(Me.txt_tahun.Text) = "" Then
    Else
        judul = judul & "Tahun " & Me.txt_tahun
    End If
    
    If Trim(Me.txt_masa.Text) = "" Then
    Else
        judul = judul & "Masa " & Me.txt_masa.Text
    End If
    
    If Trim(Me.cb_KPP) = "" Or Trim(Me.cb_KPP) = "ALL" Then
    Else
        judul = judul & "NPWP KPP " & Me.cb_KPP & ", "
    End If
    
    Me.disable_Form
    
    '-----
    'pph21tahunan2
        '0  No1, Bulan, Tahun, " & _
        '3  NPWP_KPP, kdPROYEK, kdCENTER, " & _
        '6  Nama, NPWP, NIK, " & _
        '9  Alamat, Jabatan, P_L, " & _
        '12 PTKP, Gaji, Tnj_PPh, " & _
        '15 Tunjangan_Lain, JHT_JPN, Bruto, " & _
        '18 Insentif, THR, Lainnya, " & _
        '21 Pensiun_Potongan_Lain, penghasilan_netto_sblmnya, pph21_terutang_sblmnya, " & _
        '24 nrp, id1, tglupdate, " & _
        '27 biaya_jabatan, nilai_beban " & _
    '----------
    
    jRec = RecordCount(rs)
    If jRec > 0 Then
        Call create_xls2(rs, judul, "13,14,15,16,17,18,19,20,21,22,23,27,28", "")
    End If
    Me.Enable_Form
End Sub

Private Sub cmd_Hapus_Click()
    Dim npwp_kpp As String, Tahun As String, masa As String, divisi As String, jenisPPh As String
    Dim p
    
    npwp_kpp = get_kode_combo(Me.cb_KPP, "#")
    Tahun = Me.txt_tahun
    masa = Me.txt_masa
    divisi = get_kode_combo(Me.cb_divisi, "-")
    
    If Trim(npwp_kpp) = "" Or Trim(Tahun) = "" Or Trim(masa) = "" Or Trim(divisi) = "" Then
        MsgBox "KPP harus dipilih / tahun harus di pilih / Masa harus di pilih / Divisi harus di pilih", vbInformation
        Call pesan2("batal", 1, vbYellow)
        Exit Sub
    End If
    
    p = MsgBox("Hapus data SEMUA KPP untuk " & vbCr & _
                "Tahun " & Tahun & vbCr & _
                "Masa " & masa & vbCr & _
                "Divisi " & divisi & vbCr & _
                "Proses ? ", vbYesNo)
    If p = vbYes Then
        Me.disable_Form
        Call tbPph21Tahunan2_Delete1Divisi(divisi, Tahun, masa)
        Call cmd_Load_Click
        Me.Enable_Form
    Else
        p = MsgBox("Hapus data untuk " & vbCr & _
                "KPP " & npwp_kpp & vbCr & _
                "Tahun " & Tahun & vbCr & _
                "Masa " & masa & vbCr & _
                "Divisi " & divisi & vbCr & _
                "Proses ? ", vbYesNo)
        If p = vbYes Then
            Me.disable_Form
            Call tbPph21Tahunan2_Delete1Kpp(divisi, Tahun, masa, npwp_kpp)
            Call cmd_Load_Click
            Me.Enable_Form
        End If
    End If
End Sub

Private Sub cmd_hapus1_Click()
    Dim j As Integer, rec_no As Long
    Dim nama As String, id1 As String
    Dim p
    Dim isAdaYangDihapus As Boolean
    
    '-----
    'pph21tahunan2
        '0  No1, Bulan, Tahun, " & _
        '3  NPWP_KPP, kdPROYEK, kdCENTER, " & _
        '6  Nama, NPWP, NIK, " & _
        '9  Alamat, Jabatan, P_L, " & _
        '12 PTKP, Gaji, Tnj_PPh, " & _
        '15 Tunjangan_Lain, JHT_JPN, Bruto, " & _
        '18 Insentif, THR, Lainnya, " & _
        '21 Pensiun_Potongan_Lain, penghasilan_netto_sblmnya, pph21_terutang_sblmnya, " & _
        '24 nrp, id1, tglupdate, " & _
        '27 biaya_jabatan, nilai_beban " & _
    '----------
    
    On Error GoTo er1
    isAdaYangDihapus = False
    
    If Me.DataGrid1.SelBookmarks.Count <= 0 Then
        Call pesan2("tidak ada data yang di pilih", , vbYellow)
    End If
    
    For j = 0 To Me.DataGrid1.SelBookmarks.Count - 1
        rec_no = Me.DataGrid1.SelBookmarks.Item(j)
        rs.AbsolutePosition = rec_no
        
        nama = cek_null(rs(6))
        id1 = cek_null(rs(25))
        p = MsgBox("Yakin menghapus 1 record data untuk " & vbCr & "Nama: " & nama & vbCr & _
                    "ID : " & id1 & vbCr & "?", vbYesNo)
        If p = vbYes Then
            Call tbPph21Tahunan2_Delete(id1)
            isAdaYangDihapus = True
        End If
    Next
    
    If isAdaYangDihapus = True Then Call cmd_Load_Click
    
    Exit Sub
er1:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmd_Load_Click()
    Dim sql As String
    
    Me.disable_Form
    sql = generate_sql
    DoEvents
    MsgBox sql
    If Trim(sql) <> "" Then
        Call dbMySQL_open
        If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
            MsgBox "error open sql"
            Me.Enable_Form
            Exit Sub
        End If
        
        Set Me.DataGrid1.DataSource = rs
        Call format_Grid
        Call info(1, "Jumlah data=" & RecordCount(rs), Me.StatusBar1)
    End If
    Me.Enable_Form
End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
    
  Me.txt_tahun.Text = ""
  Me.txt_masa.Text = ""
  Me.txt_cari.Text = ""
  
  '---------------------
    If dbMySQL_open = True Then
    Else
        MsgBox "Open Database Tidak Sukses", vbExclamation
    End If
    '---------------------
    
  'load combo
  Call load_Divisi(Me.cb_divisi, , 1)
    
  sql = "select distinct kdProyek from pph21tahunan2 where kdCENTER = '" & Trim(Me.cb_divisi.Text) & "'"
  Call Load_combo(Me.cb_proyek, sql, cnn, True, , 1)
  
  Call load_KPP(Me.cb_KPP, False, 1)
  
  'get level
  '2 : operator gedung
  '3 : UKP
  
  '---------
  
  Level1 = tbPengguna_getLevel1(frMenu1.nmLogin)
  If Level1 = 2 Then
    Me.cb_divisi.Text = tbPengguna_getDivisi(frMenu1.nmLogin)
    Me.cb_divisi.Enabled = False
  ElseIf Level1 = 3 Then
    Me.cb_divisi.Enabled = True
  Else
    Call pesan2("Level tidak valid", , vbYellow)
    Me.cb_divisi.Enabled = False
  End If
  
  Me.Width = 12420
    Me.Height = 7710
  
End Sub


Private Sub Form_Resize()
    With Frame3
        If Me.Width - 405 > 0 Then
            .Width = Me.Width - 405
        End If
        
        If Me.Height - 2895 > 0 Then
            .Height = Me.Height - 2895
        End If
    End With
    
    If Frame3.Height - 413 > 0 Then Me.Label6.Top = Frame3.Height - 413
    If Frame3.Height - 495 > 0 Then Me.txt_cari.Top = Frame3.Height - 495
    If Frame3.Height - 495 > 0 Then cmd_edit.Top = Frame3.Height - 495
    If Frame3.Height - 495 > 0 Then cmd_hapus1.Top = Frame3.Height - 495
    If Frame3.Height - 495 > 0 Then cmd_Hapus.Top = Frame3.Height - 495
    If Frame3.Height - 495 > 0 Then cmd_export.Top = Frame3.Height - 495
    
    With Me.DataGrid1
        If Frame3.Width - 240 > 0 Then .Width = Frame3.Width - 240
        If Frame3.Height - 840 > 0 Then .Height = Frame3.Height - 840
    End With
    
End Sub

Private Sub txt_cari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmd_Load_Click
    End If
End Sub
