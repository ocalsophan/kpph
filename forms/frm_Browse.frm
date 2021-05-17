VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Browse 
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
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "ubah data proyek, nott, nofaktur"
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmd_hapus1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Hapus Data(s)"
         Height          =   375
         Left            =   7320
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
      Begin VB.ComboBox cb_jenisPajak 
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
         Caption         =   "Jenis Pajak "
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   780
         Width           =   840
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
            Object.Width           =   10795
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10795
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
      Caption         =   "Browse Data SPT PPh & SSP PPh"
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
Attribute VB_Name = "frm_Browse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset


Function cek_Isian() As Boolean
    Dim pesan1 As String, t As String
    Dim hasil As Boolean
    
    pesan1 = ""
    hasil = True
    
    'cek divisi
    If isDataAda("mdivisi", "kodedivisi", get_kode_combo(Me.cb_divisi, "-"), cnn) = True Then
    Else
        hasil = False
        pesan1 = pesan1 & "Divisi tidak valid"
    End If
    
    'cek jenispajak
    t = get_kode_combo(Me.cb_jenisPajak, ".")
    If t = "1" Or t = "2" Or t = "3" Or t = "4" Or t = "5" Or t = "6" Or t = "7" Or t = "8" Or t = "9" Then
    Else
        hasil = False
        pesan1 = pesan1 & vbCr & "Jenis Pajak tidak valid"
    End If
    
    If Trim(pesan1) = "" Then
    Else
        MsgBox pesan1
    End If
    
    cek_Isian = hasil
End Function



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
    
    If Trim(Me.cb_divisi.text) = "ALL" Or Trim(Me.cb_divisi.text) = "" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " kode_divisi = '" & get_kode_combo(Me.cb_divisi, "-") & "'"
    End If
    
    If Trim(Me.txt_tahun) = "" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Tahun_Pajak = '" & Trim(Me.txt_tahun) & "'"
    End If
    
    If Trim(Me.txt_masa.text) = "" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Masa_Pajak = '" & Trim(Me.txt_masa.text) & "'"
    End If
    
    
    If Trim(Me.cb_kpp.text) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " NPWP_KPP = '" & get_kode_combo(Me.cb_kpp, "#") & "'"
    End If
    
    '-----
    jenisPPh = get_kode_combo(Me.cb_jenisPajak, ".")
    If Trim(jenisPPh) = "1" Then
        'pph15
        sql = "select npwp_kpp, kd_proyek, nott, " & _
                "nofaktur, kode_form, masa_pajak, " & _
                "tahun_pajak, pembetulan, npwp_wp, " & _
                "nama_wp, alamat_wp, nomor_bukti_potong, " & _
                "tanggal_bukti_potong, negara_sumber_penghasilan, kode_option_penghasilan, " & _
                "jumlah_bruto, tarif, pph_dipotong, " & _
                "invoice_ket, kode_divisi, tgl_import, " & _
                "id1 from pph15"
        If Trim(Me.txt_cari.text) <> "" Then
            cari = "kd_proyek like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nott like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nofaktur like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "npwp_wp like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nama_wp like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nomor_bukti_potong like '%" & Trim(Me.txt_cari.text) & "%' "
        End If
    ElseIf Trim(jenisPPh) = "2" Then
        'pph23
        sql = "select NPWP_KPP, kd_proyek, nott, " & _
                "nofaktur, Masa_Pajak, Kode_Form, " & _
                "Tahun_Pajak, Pembetulan, NPWP_WP, " & _
                "Nama_WP, Alamat_WP, Nomor_Bukti_Potong, " & _
                "Tanggal_Bukti_Potong, Jumlah_Nilai_Bruto_, Jumlah_PPh_Yang_Dipotong, " & _
                "kode_divisi, tgl_import, id1 from pph23 "
        If Trim(Me.txt_cari.text) <> "" Then
            cari = "kd_proyek like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nott like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nofaktur like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NPWP_WP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nama_WP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nomor_Bukti_Potong like '%" & Trim(Me.txt_cari.text) & "%' "
        End If
    ElseIf Trim(jenisPPh) = "3" Then
        'pph21tf
        sql = "select NPWP_KPP, kd_proyek, Masa_Pajak, " & _
                "Tahun_Pajak, Pembetulan, Nomor_Bukti_Potong, " & _
                "NPWP, NIK, Nama, " & _
                "Alamat, WP_Luar_Negeri, Kode_Negara, " & _
                "Kode_Pajak, Jumlah_Bruto, Jumlah_DPP, " & _
                "Tanpa_NPWP, Tarif, Jumlah_PPh, " & _
                "NPWP_Pemotong, Nama_Pemotong, Tanggal_Bukti_Potong, " & _
                "kode_divisi, tgl_import, id1 " & _
                "From pph21tf "
         If Trim(Me.txt_cari.text) <> "" Then
            cari = "kd_proyek like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NPWP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NIK like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nama like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Alamat like '%" & Trim(Me.txt_cari.text) & "%' "
         End If
    ElseIf Trim(jenisPPh) = "4" Then
        'pph21bulanan
        sql = "select NPWP_KPP, kd_proyek, Masa_Pajak, " & _
                "Tahun_Pajak, Pembetulan, NPWP, " & _
                "Nama, Kode_Pajak, Jumlah_Bruto, " & _
                "Jumlah_PPh, Kode_Negara, kode_divisi, " & _
                "tgl_import, id1 " & _
                "From pph21bulanan "
         If Trim(Me.txt_cari.text) <> "" Then
            cari = "kd_proyek like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NPWP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nama like '%" & Trim(Me.txt_cari.text) & "%'"
         End If
    ElseIf Trim(jenisPPh) = "5" Then
        'pph21tahunan
        sql = "select NPWP_KPP, kd_proyek, Masa_Pajak, " & _
                "Tahun_Pajak, Pembetulan, Nomor_Bukti_Potong, " & _
                "Masa_Perolehan_Awal, Masa_Perolehan_Akhir, NPWP, " & _
                "NIK, Nama, Alamat, " & _
                "Jenis_Kelamin, Status_PTKP, Jumlah_Tanggungan, " & _
                "Nama_Jabatan, WP_Luar_Negeri, Kode_Negara, " & _
                "Kode_Pajak, Jumlah_18, Jumlah_19, " & _
                "Jumlah_20, NPWP_Pemotong, Nama_Pemotong, " & _
                "Tanggal_Bukti_Potong, kode_divisi, tgl_import, " & _
                "id1 " & _
                "From pph21tahunan "
        If Trim(Me.txt_cari.text) <> "" Then
            cari = "kd_proyek like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NPWP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NIK like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nama like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Alamat like '%" & Trim(Me.txt_cari.text) & "%' "
         End If
    ElseIf Trim(jenisPPh) = "6" Then
        'pph22
        sql = "select NPWP_KPP, kd_proyek, nott, " & _
                "nofaktur, Masa_Pajak, Tahun_Pajak, " & _
                "Pembetulan, NPWP, Nama_NPWP, " & _
                "Alamat, Nomor_Bukti_Potong, Tanggal_Bukti_Potong, " & _
                "Nilai_DPP, Tarif, Nilai_PPh, " & _
                "j51, j52, kode_divisi, " & _
                "tgl_import, id1 " & _
                "From pph22 "
        If Trim(Me.txt_cari.text) <> "" Then
            cari = "kd_proyek like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nott like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nofaktur like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NPWP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nama_NPWP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nomor_Bukti_Potong like '%" & Trim(Me.txt_cari.text) & "%' "
        End If
    ElseIf Trim(jenisPPh) = "7" Then
        'pph26
        sql = "select NPWP_KPP, kd_proyek, nott, " & _
                "nofaktur, Masa_Pajak, Tahun_Pajak, " & _
                "Pembetulan, NPWP_WP, Nama_WP, " & _
                "Alamat_WP, Nomor_Bukti_Potong, Tanggal_Bukti_Potong, " & _
                "Jumlah_Nilai_Bruto_, Jumlah_PPh_Yang_Dipotong, kode_divisi, " & _
                "tgl_import, id1 " & _
                "From pph26 "
        If Trim(Me.txt_cari.text) <> "" Then
            cari = "kd_proyek like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nott like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nofaktur like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NPWP_WP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nama_WP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nomor_Bukti_Potong like '%" & Trim(Me.txt_cari.text) & "%' "
        End If
    ElseIf Trim(jenisPPh) = "8" Then
        'pph42_konstruksi
        sql = "select NPWP_KPP, kd_proyek, nott, " & _
                "nofaktur, Kode_Form, Masa_Pajak, " & _
                "Tahun_Pajak, Pembetulan, NPWP_WP, " & _
                "Nama_WP, Alamat_WP, Nomor_Bukti_Potong, " & _
                "Tanggal_Bukti_Potong, Jumlah_PPh_Yang_Dipotong, kode_divisi, " & _
                "tgl_import, id1,  " & _
                "Jumlah_Nilai_Bruto_1, Jumlah_Nilai_Bruto_2, Jumlah_Nilai_Bruto_3, " & _
                "Jumlah_Nilai_Bruto_4, Jumlah_Nilai_Bruto_5, Jumlah_Nilai_Bruto_6, " & _
                "Jumlah_Nilai_Bruto_7, Jumlah_Nilai_Bruto_8 " & _
                "From pph42_konstruksi "
        If Trim(Me.txt_cari.text) <> "" Then
            cari = "kd_proyek like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nott like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nofaktur like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NPWP_WP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nama_WP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nomor_Bukti_Potong like '%" & Trim(Me.txt_cari.text) & "%' "
        End If
    ElseIf Trim(jenisPPh) = "9" Then
        'pph42_sewa
        sql = "select NPWP_KPP, kd_proyek, nott, " & _
                "nofaktur, Kode_Form, Masa_Pajak, " & _
                "Tahun_Pajak, Pembetulan, NPWP_WP, " & _
                "Nama_WP, Alamat_WP, Nomor_Bukti_Potong, " & _
                "Tanggal_Bukti_Potong, Jumlah_PPh_Yang_Dipotong, kode_divisi, " & _
                "tgl_import, id1 " & _
                "From pph42_sewa "
        If Trim(Me.txt_cari.text) <> "" Then
            cari = "kd_proyek like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nott like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nofaktur like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NPWP_WP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nama_WP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nomor_Bukti_Potong like '%" & Trim(Me.txt_cari.text) & "%' "
        End If
    ElseIf Trim(jenisPPh) = "10" Then
        'pph42_sewa
        sql = "select NPWP_KPP, kd_proyek, nott, " & _
                "nofaktur, Kode_Form, Masa_Pajak, " & _
                "Tahun_Pajak, Pembetulan, NPWP_WP, " & _
                "Nama_WP, Alamat_WP, Nomor_Bukti_Potong, " & _
                "Tanggal_Bukti_Potong, Jumlah_PPh_Yang_Dipotong, kode_divisi, " & _
                "tgl_import, id1 " & _
                "From pph42_obligasi "
        If Trim(Me.txt_cari.text) <> "" Then
            cari = "kd_proyek like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nott like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "nofaktur like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NPWP_WP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nama_WP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nomor_Bukti_Potong like '%" & Trim(Me.txt_cari.text) & "%' "
        End If
    ElseIf Trim(jenisPPh) = "11" Then
        sql = "select npwp_kpp, kd_proyek, " & _
                "masa_pajak, " & _
                "tahun_pajak, pembetulan, Jumlah_karyawan, " & _
                "jumlah_bruto, " & _
                "kode_divisi, tgl_import, " & _
                "id1 from pph21_bwhptkp"
        If Trim(Me.txt_cari.text) <> "" Then
            cari = "kd_proyek like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "npwp_kpp like '%" & Trim(Me.txt_cari.text) & "%' "
        End If
    ElseIf Trim(jenisPPh) = "12" Then
        sql = "select NPWP_KPP, kd_proyek, Masa_Pajak, " & _
                "Tahun_Pajak, Pembetulan, Nomor_Bukti_Potong, " & _
                "NPWP, NIK, Nama, " & _
                "Alamat, " & _
                "Kode_Pajak, Jumlah_Bruto, " & _
                "Tarif, Jumlah_PPh, " & _
                "NPWP_Pemotong, Nama_Pemotong, Tanggal_Bukti_Potong, " & _
                "kode_divisi, tgl_import, id1 " & _
                "From pph21pesangon "
         If Trim(Me.txt_cari.text) <> "" Then
            cari = "kd_proyek like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NPWP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NIK like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Nama like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Alamat like '%" & Trim(Me.txt_cari.text) & "%' "
         End If
    ElseIf Trim(jenisPPh) = "20" Then
        'ssp_pph
        sql = "select NPWP_KPP, Jenis_Pajak, kode_divisi, " & _
                "Kode_Form, Masa_Pajak, Tahun_Pajak, " & _
                "Pembetulan, NTPN, Tanggal_Setor_SSP, " & _
                "Jumlah_SSP, Kode_KAP, Kode_Jenis_Setoran, " & _
                "tgl_import, id1 " & _
                "From ssp_pph "
        If Trim(Me.txt_cari.text) <> "" Then
            cari = "NTPN like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Jenis_Pajak like '%" & Trim(Me.txt_cari.text) & "%' "
        End If
    Else
        Call pesan2("No Reports", , vbYellow)
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
    
    generate_sql = sql
End Function

Sub format_Grid()
    
    Dim jenisPPh As String
    Dim jRec As Long
    Dim c As Integer
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
    
    '-----
    jenisPPh = get_kode_combo(Me.cb_jenisPajak, ".")
    If Trim(jenisPPh) = "1" Then
        'pph15
        '0  sql = "select npwp_kpp, kd_proyek, nott, " & _
        '3          "nofaktur, kode_form, masa_pajak, " & _
        '6         "tahun_pajak, pembetulan, npwp_wp, " & _
        '9         "nama_wp, alamat_wp, nomor_bukti_potong, " & _
        '12        "tanggal_bukti_potong, negara_sumber_penghasilan, kode_option_penghasilan, " & _
        '15        "jumlah_bruto, tarif, pph_dipotong, " & _
        '18        "invoice_ket, kode_divisi, tgl_import, " & _
        '21        "id1 from pph15"
        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
    
            If c = 12 Or c = 20 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 15 Or c = 17 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
    ElseIf Trim(jenisPPh) = "2" Then
        'pph23
        '0 sql = "select NPWP_KPP, kd_proyek, nott, " & _
        '3         "nofaktur, Masa_Pajak, Kode_Form, " & _
        '6         "Tahun_Pajak, Pembetulan, NPWP_WP, " & _
        '9         "Nama_WP, Alamat_WP, Nomor_Bukti_Potong, " & _
        '12        "Tanggal_Bukti_Potong, Jumlah_Nilai_Bruto_, Jumlah_PPh_Yang_Dipotong, " & _
        '15        "kode_divisi, tgl_import, id1 from pph23 "
        
        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
    
            If c = 12 Or c = 16 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 13 Or c = 14 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
        
    ElseIf Trim(jenisPPh) = "3" Then
        'pph21tf
        '0  sql = "select NPWP_KPP, kd_proyek, Masa_Pajak, " & _
        '3         "Tahun_Pajak, Pembetulan, Nomor_Bukti_Potong, " & _
        '6        "NPWP, NIK, Nama, " & _
        '9        "Alamat, WP_Luar_Negeri, Kode_Negara, " & _
        '12        "Kode_Pajak, Jumlah_Bruto, Jumlah_DPP, " & _
        '15        "Tanpa_NPWP, Tarif, Jumlah_PPh, " & _
        '18        "NPWP_Pemotong, Nama_Pemotong, Tanggal_Bukti_Potong, " & _
        '21        "kode_divisi, tgl_import, id1 " & _

        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
    
            If c = 20 Or c = 22 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 13 Or c = 14 Or c = 17 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
    ElseIf Trim(jenisPPh) = "4" Then
        'pph21bulanan
        '0 sql = "select NPWP_KPP, kd_proyek, Masa_Pajak, " & _
        '3        "Tahun_Pajak, Pembetulan, NPWP, " & _
        '6        "Nama, Kode_Pajak, Jumlah_Bruto, " & _
        '9        "Jumlah_PPh, Kode_Negara, kode_divisi, " & _
        '12        "tgl_import, id1 " & _

        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
    
            If c = 12 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 8 Or c = 9 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next

    ElseIf Trim(jenisPPh) = "5" Then
        'pph21tahunan
        '0 sql = "select NPWP_KPP, kd_proyek, Masa_Pajak, " & _
        '3        "Tahun_Pajak, Pembetulan, Nomor_Bukti_Potong, " & _
        '6        "Masa_Perolehan_Awal, Masa_Perolehan_Akhir, NPWP, " & _
        '9        "NIK, Nama, Alamat, " & _
        '12        "Jenis_Kelamin, Status_PTKP, Jumlah_Tanggungan, " & _
        '15        "Nama_Jabatan, WP_Luar_Negeri, Kode_Negara, " & _
        '18        "Kode_Pajak, Jumlah_18, Jumlah_19, " & _
        '21        "Jumlah_20, NPWP_Pemotong, Nama_Pemotong, " & _
        '24        "Tanggal_Bukti_Potong, kode_divisi, tgl_import, " & _
        '27        "id1 " & _

        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
    
            If c = 24 Or c = 26 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 19 Or c = 20 Or c = 21 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
    ElseIf Trim(jenisPPh) = "6" Then
        'pph22
        '0 sql = "select NPWP_KPP, kd_proyek, nott, " & _
        '3        "nofaktur, Masa_Pajak, Tahun_Pajak, " & _
        '6        "Pembetulan, NPWP, Nama_NPWP, " & _
        '9        "Alamat, Nomor_Bukti_Potong, Tanggal_Bukti_Potong, " & _
        '12        "Nilai_DPP, Tarif, Nilai_PPh, " & _
        '15        "j51, j52, kode_divisi, " & _
        '18        "tgl_import, id1 " & _

        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
    
            If c = 11 Or c = 18 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 12 Or c = 14 Or c = 15 Or c = 16 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
        
    ElseIf Trim(jenisPPh) = "7" Then
        'pph26
        '0 sql = "select NPWP_KPP, kd_proyek, nott, " & _
        '3        "nofaktur, Masa_Pajak, Tahun_Pajak, " & _
        '6        "Pembetulan, NPWP_WP, Nama_WP, " & _
        '9        "Alamat_WP, Nomor_Bukti_Potong, Tanggal_Bukti_Potong, " & _
        '12        "Jumlah_Nilai_Bruto_, Jumlah_PPh_Yang_Dipotong, kode_divisi, " & _
        '15        "tgl_import, id1 " & _

        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
    
            If c = 11 Or c = 15 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 12 Or c = 13 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next

    ElseIf Trim(jenisPPh) = "8" Then
        'pph42_konstruksi
        '0 sql = "select NPWP_KPP, kd_proyek, nott, " & _
        '3        "nofaktur, Kode_Form, Masa_Pajak, " & _
        '6        "Tahun_Pajak, Pembetulan, NPWP_WP, " & _
        '9        "Nama_WP, Alamat_WP, Nomor_Bukti_Potong, " & _
        '12        "Tanggal_Bukti_Potong, Jumlah_PPh_Yang_Dipotong, kode_divisi, " & _
        '15        "tgl_import, id1 " & _
        '17         bruto1, 2, 3, 4, 5, 6, 7, 8

        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
    
            If c = 12 Or c = 15 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 13 Or c = 17 Or c = 18 Or c = 19 Or c = 20 Or c = 21 Or c = 22 Or c = 23 Or c = 24 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
        
    ElseIf Trim(jenisPPh) = "9" Then
        'pph42_sewa
        '0 sql = "select NPWP_KPP, kd_proyek, nott, " & _
        '3        "nofaktur, Kode_Form, Masa_Pajak, " & _
        '6        "Tahun_Pajak, Pembetulan, NPWP_WP, " & _
        '9        "Nama_WP, Alamat_WP, Nomor_Bukti_Potong, " & _
        '12        "Tanggal_Bukti_Potong, Jumlah_PPh_Yang_Dipotong, kode_divisi, " & _
        '15        "tgl_import, id1 " & _

        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
    
            If c = 12 Or c = 15 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 13 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
    ElseIf Trim(jenisPPh) = "10" Then

        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
    
            If c = 12 Or c = 15 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 13 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
    ElseIf Trim(jenisPPh) = "11" Then
        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
    
            If c = 12 Or c = 20 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 6 Or c = 17 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
    ElseIf Trim(jenisPPh) = "12" Then

        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
    
            If c = 20 Or c = 22 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 13 Or c = 14 Or c = 17 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
    ElseIf Trim(jenisPPh) = "20" Then
        'ssp_pph
        '0 sql = "select NPWP_KPP, Jenis_Pajak, kode_divisi, " & _
        '3        "Kode_Form, Masa_Pajak, Tahun_Pajak, " & _
        '6        "Pembetulan, NTPN, Tanggal_Setor_SSP, " & _
        '9        "Jumlah_SSP, Kode_KAP, Kode_Jenis_Setoran, " & _
        '12        "tgl_import, id1 " & _
        '15        "From ssp_pph "
        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
    
            If c = 8 Or c = 12 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
                Me.DataGrid1.Columns(c).Width = 900
            End If
    
            If c = 9 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
    Else
        Call pesan2("No Reports", , vbYellow)
    End If
End Sub

Private Sub cb_kpp_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 113 And Shift = 0 Then Call load_KPP(Me.cb_kpp, True)
End Sub

Private Sub cmd_edit_Click()
    Dim kd_proyek As String, nott As String, nofaktur As String, jenisPPh As String, id1 As String
    Dim jRec As Long
    Dim indexAkhir As Integer
    
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
    kd_proyek = cek_null(rs(1))
    nott = cek_null(rs(2))
    nofaktur = cek_null(rs(3))
    
    kd_proyek = InputBox("kd proyek", "edit", kd_proyek)
    nott = InputBox("noTT", "edit", nott)
    nofaktur = InputBox("noFaktur", "edit", nofaktur)
    
    
    indexAkhir = rs.Fields.Count - 1
    id1 = cek_null(rs(indexAkhir))
    
    jenisPPh = get_kode_combo(Me.cb_jenisPajak, ".")
            If Trim(jenisPPh) = "1" Then
                'pph15
                Call tbPphX_editById(id1, "pph15", nott, nofaktur, kd_proyek)
            ElseIf Trim(jenisPPh) = "2" Then
                'pph23
                Call tbPphX_editById(id1, "pph23", nott, nofaktur, kd_proyek)
            ElseIf Trim(jenisPPh) = "3" Then
                'pph21tf
                Call tbPphX_editById(id1, "pph21tf", nott, nofaktur, kd_proyek)
            ElseIf Trim(jenisPPh) = "4" Then
                'pph21bulanan
                Call tbPphX_editById(id1, "pph21bulanan", nott, nofaktur, kd_proyek)
            ElseIf Trim(jenisPPh) = "5" Then
                'pph21tahunan
                Call tbPphX_editById(id1, "pph21tahunan", nott, nofaktur, kd_proyek)
            ElseIf Trim(jenisPPh) = "6" Then
                'pph22
                Call tbPphX_editById(id1, "pph22", nott, nofaktur, kd_proyek)
            ElseIf Trim(jenisPPh) = "7" Then
                'pph26
                Call tbPphX_editById(id1, "pph26", nott, nofaktur, kd_proyek)
            ElseIf Trim(jenisPPh) = "8" Then
                'pph42_konstruksi
                Call tbPphX_editById(id1, "pph42_konstruksi", nott, nofaktur, kd_proyek)
            ElseIf Trim(jenisPPh) = "9" Then
                'pph42_sewa
                Call tbPphX_editById(id1, "pph42_sewa", nott, nofaktur, kd_proyek)
            Else
                Call pesan2("No Reports", , vbYellow)
            End If
    Call cmd_load_Click
End Sub

Private Sub cmd_export_Click()
    Dim jRec As Long
    
    Me.disable_Form
    jRec = RecordCount(rs)
    If jRec > 0 Then
        Call create_xls2(rs, Me.cb_jenisPajak, "", "")
    End If
    Me.Enable_Form
End Sub

Private Sub cmd_hapus_Click()
    Dim npwp_kpp As String, tahun As String, masa As String, DIVISI As String, jenisPPh As String
    Dim p
    
    npwp_kpp = Me.cb_kpp
    tahun = Me.txt_tahun
    masa = Me.txt_masa
    DIVISI = Me.cb_divisi
    
    If Trim(npwp_kpp) = "" Or Trim(tahun) = "" Or Trim(masa) = "" Or Trim(DIVISI) = "" Then
        MsgBox "KPP harus dipilih / tahun harus di pilih / Masa harus di pilih / Divisi harus di pilih", vbInformation
        Call pesan2("batal", 1, vbYellow)
        Exit Sub
    End If
    
    p = MsgBox("Hapus data untuk " & vbCr & _
                "KPP " & npwp_kpp & vbCr & _
                "Tahun " & tahun & vbCr & _
                "Masa " & masa & vbCr & _
                "Divisi " & DIVISI & vbCr & _
                "Proses ? ", vbYesNo)
    If p = vbYes Then
        Me.disable_Form
        
        jenisPPh = get_kode_combo(Me.cb_jenisPajak, ".")
            If Trim(jenisPPh) = "1" Then
                'pph15
                Call tbPphX_deleteByKPP("pph15", get_kode_combo(Me.cb_kpp, "#"), tahun, masa, get_kode_combo(Me.cb_divisi, "-"))
            ElseIf Trim(jenisPPh) = "2" Then
                'pph23
                Call tbPphX_deleteByKPP("pph23", get_kode_combo(Me.cb_kpp, "#"), tahun, masa, get_kode_combo(Me.cb_divisi, "-"))
            ElseIf Trim(jenisPPh) = "3" Then
                'pph21tf
                Call tbPphX_deleteByKPP("pph21tf", get_kode_combo(Me.cb_kpp, "#"), tahun, masa, get_kode_combo(Me.cb_divisi, "-"))
            ElseIf Trim(jenisPPh) = "4" Then
                'pph21bulanan
                Call tbPphX_deleteByKPP("pph21bulanan", get_kode_combo(Me.cb_kpp, "#"), tahun, masa, get_kode_combo(Me.cb_divisi, "-"))
            ElseIf Trim(jenisPPh) = "5" Then
                'pph21tahunan
                Call tbPphX_deleteByKPP("pph21tahunan", get_kode_combo(Me.cb_kpp, "#"), tahun, masa, get_kode_combo(Me.cb_divisi, "-"))
            ElseIf Trim(jenisPPh) = "6" Then
                'pph22
                Call tbPphX_deleteByKPP("pph22", get_kode_combo(Me.cb_kpp, "#"), tahun, masa, get_kode_combo(Me.cb_divisi, "-"))
            ElseIf Trim(jenisPPh) = "7" Then
                'pph26
                Call tbPphX_deleteByKPP("pph26", get_kode_combo(Me.cb_kpp, "#"), tahun, masa, get_kode_combo(Me.cb_divisi, "-"))
            ElseIf Trim(jenisPPh) = "8" Then
                'pph42_konstruksi
                Call tbPphX_deleteByKPP("pph42_konstruksi", get_kode_combo(Me.cb_kpp, "#"), tahun, masa, get_kode_combo(Me.cb_divisi, "-"))
            ElseIf Trim(jenisPPh) = "9" Then
                'pph42_sewa
                Call tbPphX_deleteByKPP("pph42_sewa", get_kode_combo(Me.cb_kpp, "#"), tahun, masa, get_kode_combo(Me.cb_divisi, "-"))
            ElseIf Trim(jenisPPh) = "10" Then
                'ssp_pph
                Call tbPphX_deleteByKPP("ssp_pph", get_kode_combo(Me.cb_kpp, "#"), tahun, masa, get_kode_combo(Me.cb_divisi, "-"))
            Else
                Call pesan2("No Reports", , vbYellow)
            End If
            
        Call cmd_load_Click
        Me.Enable_Form
    End If
End Sub

Private Sub cmd_hapus1_Click()
    Dim j As Integer, rec_no As Long
    Dim npwp_kpp As String, id1 As String, jenisPPh As String
    Dim p
    Dim indexAkhir As Integer
    Dim isAdaYangDihapus As Boolean
    
    On Error GoTo er1
    isAdaYangDihapus = False
    indexAkhir = rs.Fields.Count - 1
    For j = 0 To Me.DataGrid1.SelBookmarks.Count - 1
        rec_no = Me.DataGrid1.SelBookmarks.Item(j)
        rs.AbsolutePosition = rec_no
        
        npwp_kpp = cek_null(rs(0))
        id1 = cek_null(rs(indexAkhir))
        p = MsgBox("Yakin menghapus 1 record data untuk " & vbCr & "KPP: " & npwp_kpp & vbCr & _
                    "ID : " & id1 & vbCr & "?", vbYesNo)
        If p = vbYes Then
            isAdaYangDihapus = True
            jenisPPh = get_kode_combo(Me.cb_jenisPajak, ".")
            If Trim(jenisPPh) = "1" Then
                'pph15
                Call tbPphX_deleteById(id1, "pph15")
            ElseIf Trim(jenisPPh) = "2" Then
                'pph23
                Call tbPphX_deleteById(id1, "pph23")
            ElseIf Trim(jenisPPh) = "3" Then
                'pph21tf
                Call tbPphX_deleteById(id1, "pph21tf")
            ElseIf Trim(jenisPPh) = "4" Then
                'pph21bulanan
                Call tbPphX_deleteById(id1, "pph21bulanan")
            ElseIf Trim(jenisPPh) = "5" Then
                'pph21tahunan
                Call tbPphX_deleteById(id1, "pph21tahunan")
            ElseIf Trim(jenisPPh) = "6" Then
                'pph22
                Call tbPphX_deleteById(id1, "pph22")
            ElseIf Trim(jenisPPh) = "7" Then
                'pph26
                Call tbPphX_deleteById(id1, "pph26")
            ElseIf Trim(jenisPPh) = "8" Then
                'pph42_konstruksi
                Call tbPphX_deleteById(id1, "pph42_konstruksi")
            ElseIf Trim(jenisPPh) = "9" Then
                'pph42_sewa
                Call tbPphX_deleteById(id1, "pph42_sewa")
            ElseIf Trim(jenisPPh) = "10" Then
                'pph42_sewa
                Call tbPphX_deleteById(id1, "ssp_pph")
            Else
                Call pesan2("No Reports", , vbYellow)
            End If
        End If
    Next
    
    If isAdaYangDihapus = True Then Call cmd_load_Click
    
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
End Sub

Private Sub cmd_load_Click()
    Dim sql As String, jRec As Long
    
    Me.disable_Form
    sql = generate_sql
    DoEvents
    'MsgBox sql
    If Trim(sql) <> "" Then
        Call dbMySQL_open
        If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
            MsgBox "error open sql"
            Me.Enable_Form
            Exit Sub
        End If
        
        Set Me.DataGrid1.DataSource = rs
        jRec = RecordCount(rs)
        Call format_Grid
        Call info(1, "Jumlah data : " & jRec, Me.StatusBar1)
    End If
    Me.Enable_Form
End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
  Call dbMySQL_open
    
  Me.txt_tahun.text = ""
  Me.txt_masa.text = ""
  Me.txt_cari.text = ""
    
  'load combo
  Call load_Divisi(Me.cb_divisi, , 1)
  
  Call load_jenisPPh(Me.cb_jenisPajak)
  cb_jenisPajak.AddItem "20. SSP PPH"
  
  Call load_KPP(Me.cb_kpp, False, 1)
  
  'get level
  '2 : operator gedung
  '3 : UKP
  
  '---------
  
  Level1 = tbPengguna_getLevel1(frMenu1.nmLogin)
  If Level1 = 2 Then
    Me.cb_divisi.text = tbPengguna_getDivisi(frMenu1.nmLogin)
    Me.cb_divisi.Enabled = False
  ElseIf Level1 = 3 Then
    Me.cb_divisi.Enabled = True
  Else
    Call pesan2("Level tidak valid", , vbYellow)
    Me.cb_divisi.Enabled = False
  End If
  
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call dbMySQL_close
End Sub

Private Sub txt_cari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmd_load_Click
    End If
End Sub
