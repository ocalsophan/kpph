VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_csv_SPT 
   ClientHeight    =   7590
   ClientLeft      =   300
   ClientTop       =   810
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
   ScaleHeight     =   7590
   ScaleWidth      =   12300
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4455
      Left            =   240
      TabIndex        =   18
      Top             =   2760
      Width           =   11895
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3615
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   6376
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
      Begin VB.CommandButton cmd_export 
         Cancel          =   -1  'True
         Caption         =   "Export CSV"
         Height          =   375
         Left            =   10200
         TabIndex        =   9
         Top             =   3960
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmd_proses 
      Caption         =   "Load"
      Height          =   375
      Left            =   10560
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   " 1. Divisi / Jenis PPh / KPP "
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
      Width           =   12015
      Begin VB.OptionButton opt2 
         Caption         =   "Format 2 (KPP + Divisi)"
         Height          =   375
         Left            =   9600
         TabIndex        =   8
         Top             =   960
         Width           =   2300
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Format 1"
         Height          =   375
         Left            =   9600
         TabIndex        =   20
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cb_pembetulan 
         Height          =   330
         Left            =   7800
         TabIndex        =   6
         Text            =   "x"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox cb_masa 
         Height          =   330
         Left            =   7800
         TabIndex        =   5
         Text            =   "x"
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cb_tahun 
         Height          =   330
         Left            =   7800
         TabIndex        =   4
         Text            =   "x"
         Top             =   300
         Width           =   1575
      End
      Begin VB.ComboBox cb_KPP 
         Height          =   330
         Left            =   1200
         TabIndex        =   3
         Text            =   "Combo1"
         ToolTipText     =   "F2 untuk Filter"
         Top             =   1080
         Width           =   5535
      End
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
         Top             =   300
         Width           =   2655
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Format Output"
         Height          =   210
         Left            =   9600
         TabIndex        =   19
         Top             =   360
         Width           =   1020
      End
      Begin VB.Line Line3 
         X1              =   9480
         X2              =   9480
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Pembetulan"
         Height          =   210
         Left            =   6960
         TabIndex        =   17
         Top             =   1140
         Width           =   825
      End
      Begin VB.Line Line1 
         X1              =   6840
         X2              =   6840
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Masa"
         Height          =   210
         Left            =   6960
         TabIndex        =   16
         Top             =   780
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   6960
         TabIndex        =   15
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "KPP"
         Height          =   210
         Left            =   240
         TabIndex        =   14
         Top             =   1140
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pajak "
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   780
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Divisi"
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   7335
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
      Caption         =   "CSV SPT PPh"
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
Attribute VB_Name = "frm_csv_SPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsGrid As ADODB.Recordset

Function cek_Isian() As Boolean
    Dim pesan1 As String, t As String
    Dim hasil As Boolean
    
    pesan1 = ""
    hasil = True
    
    'cek divisi
    If Trim(Me.cb_divisi.text) = "" Then
        hasil = False
        pesan1 = pesan1 & "Divisi tidak valid"
    End If
    
    'cek jenispajak
    t = get_kode_combo(Me.cb_jenisPajak, ".")
    If t = "1" Or t = "2" Or t = "3" Or t = "4" Or t = "5" Or t = "6" Or _
        t = "7" Or t = "8" Or t = "9" Or t = "10" Or t = "11" Or t = "12" Then
    Else
        hasil = False
        pesan1 = pesan1 & vbCr & "Jenis Pajak tidak valid"
    End If
    
    'cek KPP
    If Trim(Me.cb_kpp.text) = "" Then
        hasil = False
        pesan1 = pesan1 & vbCr & "KPP tidak valid"
    End If
    
    If Trim(pesan1) = "" Then
    Else
        MsgBox pesan1
    End If
    
    cek_Isian = hasil
End Function

Private Sub cb_jenisPajak_Click()
    Dim jenisPPh As String
    
    Me.disable_Form
    Call dbMySQL_open
    
    jenisPPh = get_kode_combo(Me.cb_jenisPajak, ".")
    If Trim(jenisPPh) = "1" Then
        Call load_Tahun_pph15(Me.cb_tahun)
        Call load_Masa_pph15(Me.cb_masa)
        Call load_Pembetulan2(Me.cb_pembetulan, "pph15")
    ElseIf Trim(jenisPPh) = "2" Then
        Call load_Tahun2(Me.cb_tahun, "pph23")
        Call load_Masa2(Me.cb_masa, "pph23")
        Call load_Pembetulan2(Me.cb_pembetulan, "pph23")
    ElseIf Trim(jenisPPh) = "3" Then
        Call load_Tahun2(Me.cb_tahun, "pph21tf")
        Call load_Masa2(Me.cb_masa, "pph21tf")
        Call load_Pembetulan2(Me.cb_pembetulan, "pph21tf")
    ElseIf Trim(jenisPPh) = "4" Then
        Call load_Tahun2(Me.cb_tahun, "pph21bulanan")
        Call load_Masa2(Me.cb_masa, "pph21bulanan")
        Call load_Pembetulan2(Me.cb_pembetulan, "pph21bulanan")
    ElseIf Trim(jenisPPh) = "5" Then
        Call load_Tahun2(Me.cb_tahun, "pph21tahunan")
        Call load_Masa2(Me.cb_masa, "pph21tahunan")
        Call load_Pembetulan2(Me.cb_pembetulan, "pph21tahunan")
    ElseIf Trim(jenisPPh) = "6" Then
        Call load_Tahun2(Me.cb_tahun, "pph22")
        Call load_Masa2(Me.cb_masa, "pph22")
        Call load_Pembetulan2(Me.cb_pembetulan, "pph22")
    ElseIf Trim(jenisPPh) = "7" Then
        Call load_Tahun2(Me.cb_tahun, "pph26")
        Call load_Masa2(Me.cb_masa, "pph26")
        Call load_Pembetulan2(Me.cb_pembetulan, "pph26")
    ElseIf Trim(jenisPPh) = "8" Then
        Call load_Tahun2(Me.cb_tahun, "pph42_konstruksi")
        Call load_Masa2(Me.cb_masa, "pph42_konstruksi")
        Call load_Pembetulan2(Me.cb_pembetulan, "pph42_konstruksi")
    ElseIf Trim(jenisPPh) = "9" Then
        Call load_Tahun2(Me.cb_tahun, "pph42_sewa")
        Call load_Masa2(Me.cb_masa, "pph42_sewa")
        Call load_Pembetulan2(Me.cb_pembetulan, "pph42_sewa")
    ElseIf Trim(jenisPPh) = "10" Then
        Call load_Tahun2(Me.cb_tahun, "pph42_obligasi")
        Call load_Masa2(Me.cb_masa, "pph42_obligasi")
        Call load_Pembetulan2(Me.cb_pembetulan, "pph42_obligasi")
    ElseIf Trim(jenisPPh) = "11" Then
        Call load_Tahun2(Me.cb_tahun, "pph21_bwhptkp")
        Call load_Masa2(Me.cb_masa, "pph21_bwhptkp")
        Call load_Pembetulan2(Me.cb_pembetulan, "pph21_bwhptkp")
    ElseIf Trim(jenisPPh) = "12" Then
        Call load_Tahun2(Me.cb_tahun, "pph21pesangon")
        Call load_Masa2(Me.cb_masa, "pph21pesangon")
        Call load_Pembetulan2(Me.cb_pembetulan, "pph21pesangon")
    Else
        Me.cb_tahun.Clear
        Me.cb_masa.Clear
    End If
    Me.Enable_Form
End Sub

Private Sub cb_kpp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 And Shift = 0 Then Call load_KPP(Me.cb_kpp, True)
End Sub




Sub disable_Form()
    Me.Frame1.Enabled = False
    Me.Frame2.Enabled = False
    Me.cmd_proses.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame1.Enabled = True
    Me.Frame2.Enabled = True
    Me.cmd_proses.Enabled = True
End Sub


Private Sub cmd_export_Click()
    Dim nmFile As String
    Dim jenisPPh As String, nmPPh As String
    
    jenisPPh = get_kode_combo(Me.cb_jenisPajak, ".")
    If Trim(jenisPPh) = "1" Then
        nmPPh = "pph15"
    ElseIf Trim(jenisPPh) = "2" Then
        nmPPh = "pph23"
    ElseIf Trim(jenisPPh) = "3" Then
        nmPPh = "pph21TidakFinal"
    ElseIf Trim(jenisPPh) = "4" Then
        nmPPh = "pph21Bulanan"
    ElseIf Trim(jenisPPh) = "5" Then
        nmPPh = "pph21Tahunan"
    ElseIf Trim(jenisPPh) = "6" Then
        nmPPh = "pph22"
    ElseIf Trim(jenisPPh) = "7" Then
        nmPPh = "pph26"
    ElseIf Trim(jenisPPh) = "8" Then
        nmPPh = "pph42_konstruksi"
    ElseIf Trim(jenisPPh) = "9" Then
        nmPPh = "pph42_sewa"
    Else
        nmPPh = "err"
    End If
    
    nmFile = App.Path & "\exp\" & getTimeStamp(Now) & "_" & nmPPh & "_" & get_kode_combo(Me.cb_kpp, "#") & "_" & Trim(Me.cb_tahun) & _
            "_" & Trim(Me.cb_masa) & ".csv"

    If Trim(jenisPPh) = "1" Then
        Call create_csv(rsGrid, nmFile, , False, , "", "09")
    ElseIf Trim(jenisPPh) = "3" Then
        Call create_csv(rsGrid, nmFile, , False, , "", "09")
    ElseIf Trim(jenisPPh) = "4" Then
        Call create_csv(rsGrid, nmFile, , False, , "", "08")
    ElseIf Trim(jenisPPh) = "5" Then
        Call create_csv(rsGrid, nmFile, , False, , "", "15")
    Else
        Call create_csv(rsGrid, nmFile, , False, , "")
    End If
    
    
End Sub

Private Sub cmd_proses_Click()
    Dim jenisPPh As String
    
    On Error GoTo er1
    
    Me.disable_Form
    
    '---------------------
    If dbMySQL_open = True Then
    Else
        MsgBox "Open Database Tidak Sukses", vbExclamation
        Me.Enable_Form
        Exit Sub
    End If
    '---------------------
    
    If cek_Isian() = False Then
        Me.Enable_Form
        Exit Sub
    End If
        
    jenisPPh = get_kode_combo(Me.cb_jenisPajak, ".")
    If Trim(jenisPPh) = "1" Then
    
        Call load_data_Csv("pph15", 4, 18, _
                                "Kode Form Bukti Potong / Kode Form Input PPh Yang Dibayar Sendiri;Masa Pajak;Tahun Pajak;Pembetulan;NPWP WP yang Dipotong; Nama WP yang Dipotong;Alamat WP yang Dipotong;Nomor Bukti Potong / Nomor Urut Pada PPh Pasal 24 Yang Dapat Diperhitungkan / NTPP;Tanggal Bukti Potong / Tanggal SSP;Negara Sumber Penghasilan;Kode Option Penghasilan;Jumlah Bruto / Jumlah Penghasilan Pada Form Input Yang Dibayar Sendiri;Tarif  /  Jumlah Pajak Terutang yang dibayar di luar negeri;PPh Yang Dipotong  /  PPh Pasal 24 Yang Dapat Diperhitungkan / Jumlah PPh Pada Form Input Yang Dibayar Sendiri;Invoice / Keterangan", _
                                Me.StatusBar1)
    ElseIf Trim(jenisPPh) = "2" Then
            
        Call load_data_Csv("pph23", 4, 80, _
   "Kode Form Bukti Potong;Masa Pajak;Tahun Pajak;Pembetulan;NPWP WP yang Dipotong;Nama WP yang Dipotong;Alamat WP yang Dipotong;Nomor Bukti Potong;Tanggal Bukti Potong;Nilai Bruto 1;Tarif 1;PPh Yang Dipotong  1;Nilai Bruto 2;Tarif 2;PPh Yang Dipotong  2;Nilai Bruto 3;Tarif 3;PPh Yang Dipotong  3;Nilai Bruto 4;Tarif 4;PPh Yang Dipotong  4;Nilai Bruto 5;Tarif 5;PPh Yang Dipotong  5;Nilai Bruto 6a/Nilai Bruto 6;Tarif 6a/Tarif 6;PPh Yang Dipotong  6a/PPh Yang Dipotong  6;Nilai Bruto 6b/Nilai Bruto 7;Tarif 6b/Tarif 7;PPh Yang Dipotong  6b/PPh Yang Dipotong  7;Nilai Bruto 6c/Nilai Bruto 8;Tarif 6c/Tarif 8;PPh Yang Dipotong  6c/PPh Yang Dipotong  8;Nilai Bruto 9;Tarif 9;PPh Yang Dipotong  9;Nilai Bruto 10;Perkiraan Penghasilan Netto10;Tarif 10;PPh Yang Dipotong  10;Nilai Bruto 11;Perkiraan Penghasilan Netto11;Tarif 11;PPh Yang Dipotong  11;Nilai Bruto 12;Perkiraan Penghasilan Netto12;Tarif 12;PPh Yang Dipotong  12;Nilai Bruto 13;Tarif 13;PPh Yang Dipotong  13;Kode Jasa 6d1 PMK-244/PMK.03/2008;Nilai Bruto 6d1;" & _
    "Tarif 6d1;PPh Yang Dipotong  6d1;Kode Jasa 6d2 PMK-244/PMK.03/2008;Nilai Bruto 6d2;Tarif 6d2;PPh Yang Dipotong  6d2;Kode Jasa 6d3 PMK-244/PMK.03/2008;Nilai Bruto 6d3;Tarif 6d3;PPh Yang Dipotong  6d3;Kode Jasa 6d4 PMK-244/PMK.03/2008;Nilai Bruto 6d4;Tarif 6d4;PPh Yang Dipotong  6d4;Kode Jasa 6d5 PMK-244/PMK.03/2008;Nilai Bruto 6d5;Tarif 6d5;PPh Yang Dipotong  6d5;Kode Jasa 6d6 PMK-244/PMK.03/2008;Nilai Bruto 6d6;Tarif 6d6;PPh Yang Dipotong  6d6;Jumlah Nilai Bruto ;Jumlah PPh Yang Dipotong", _
                                Me.StatusBar1)
    
    ElseIf Trim(jenisPPh) = "3" Then
    
        Call load_data_Csv("pph21tf", 2, 20, _
                                "Masa Pajak;Tahun Pajak;Pembetulan;Nomor Bukti Potong;NPWP;NIK;Nama;Alamat;WP Luar Negeri;Kode Negara;Kode Pajak;Jumlah Bruto;Jumlah DPP;Tanpa NPWP;Tarif;Jumlah PPh;NPWP Pemotong;Nama Pemotong;Tanggal Bukti Potong", _
                                Me.StatusBar1)
                                
    ElseIf Trim(jenisPPh) = "4" Then
    
        Call load_data_Csv("pph21bulanan", 2, 10, _
                                "Masa Pajak;Tahun Pajak;Pembetulan;NPWP;Nama;Kode Pajak;Jumlah Bruto;Jumlah PPh;Kode Negara", _
                                Me.StatusBar1)
                                
    ElseIf Trim(jenisPPh) = "5" Then
    
        Call load_data_Csv("pph21tahunan", 2, 42, _
                                "Masa Pajak;Tahun Pajak;Pembetulan;Nomor Bukti Potong;Masa Perolehan Awal;Masa Perolehan Akhir;NPWP;NIK;Nama;Alamat;Jenis Kelamin;Status PTKP;Jumlah Tanggungan;Nama Jabatan;WP Luar Negeri;Kode Negara;Kode Pajak;Jumlah 1;Jumlah 2;Jumlah 3;Jumlah 4;Jumlah 5;Jumlah 6;Jumlah 7;Jumlah 8;Jumlah 9;Jumlah 10;Jumlah 11;Jumlah 12;Jumlah 13;Jumlah 14;Jumlah 15;Jumlah 16;Jumlah 17;Jumlah 18;Jumlah 19;Jumlah 20;Status Pindah;NPWP Pemotong;Nama Pemotong;Tanggal Bukti Potong", _
                                Me.StatusBar1)
                                
    ElseIf Trim(jenisPPh) = "6" Then
    
        Call load_data_Csv("pph22", 4, 54, _
                                "", _
                                Me.StatusBar1)
                                
    ElseIf Trim(jenisPPh) = "7" Then
        
        Call load_data_Csv("pph26", 4, 80, _
   "Kode Form Bukti Potong;Masa Pajak;Tahun Pajak;Pembetulan;NPWP WP yang Dipotong;Nama WP yang Dipotong;Alamat WP yang Dipotong;Nomor Bukti Potong;Tanggal Bukti Potong;Nilai Bruto 1;Tarif 1;PPh Yang Dipotong  1;Nilai Bruto 2;Tarif 2;PPh Yang Dipotong  2;Nilai Bruto 3;Tarif 3;PPh Yang Dipotong  3;Nilai Bruto 4;Tarif 4;PPh Yang Dipotong  4;Nilai Bruto 5;Tarif 5;PPh Yang Dipotong  5;Nilai Bruto 6a/Nilai Bruto 6;Tarif 6a/Tarif 6;PPh Yang Dipotong  6a/PPh Yang Dipotong  6;Nilai Bruto 6b/Nilai Bruto 7;Tarif 6b/Tarif 7;PPh Yang Dipotong  6b/PPh Yang Dipotong  7;Nilai Bruto 6c/Nilai Bruto 8;Tarif 6c/Tarif 8;PPh Yang Dipotong  6c/PPh Yang Dipotong  8;Nilai Bruto 9;Tarif 9;PPh Yang Dipotong  9;Nilai Bruto 10;Perkiraan Penghasilan Netto10;Tarif 10;PPh Yang Dipotong  10;Nilai Bruto 11;Perkiraan Penghasilan Netto11;Tarif 11;PPh Yang Dipotong  11;Nilai Bruto 12;Perkiraan Penghasilan Netto12;Tarif 12;PPh Yang Dipotong  12;Nilai Bruto 13;Tarif 13;PPh Yang Dipotong  13;Kode Jasa 6d1 PMK-244/PMK.03/2008;Nilai Bruto 6d1;" & _
    "Tarif 6d1;PPh Yang Dipotong  6d1;Kode Jasa 6d2 PMK-244/PMK.03/2008;Nilai Bruto 6d2;Tarif 6d2;PPh Yang Dipotong  6d2;Kode Jasa 6d3 PMK-244/PMK.03/2008;Nilai Bruto 6d3;Tarif 6d3;PPh Yang Dipotong  6d3;Kode Jasa 6d4 PMK-244/PMK.03/2008;Nilai Bruto 6d4;Tarif 6d4;PPh Yang Dipotong  6d4;Kode Jasa 6d5 PMK-244/PMK.03/2008;Nilai Bruto 6d5;Tarif 6d5;PPh Yang Dipotong  6d5;Kode Jasa 6d6 PMK-244/PMK.03/2008;Nilai Bruto 6d6;Tarif 6d6;PPh Yang Dipotong  6d6;Jumlah Nilai Bruto ;Jumlah PPh Yang Dipotong", _
                                Me.StatusBar1)
        
    ElseIf Trim(jenisPPh) = "8" Then
    
        Call load_data_Csv("pph42_konstruksi", 4, 54, _
                            "NPWP KPP;Kode Form Bukti Potong / Kode Form Input PPh Yang Dibayar Sendiri;Masa Pajak;Tahun Pajak;Pembetulan;" & _
                            "NPWP WP yang Dipotong;Nama WP yang Dipotong;Alamat WP yang Dipotong;Nomor Bukti Potong / NTPN;Tanggal Bukti Potong/Tanggal SSP;Jenis Hadiah Undian 1 / Lokasi Tanah dan atau Bangunan / Nama Obligasi;Kode Option Tempat Penyimpanan 1 (Khusus F113310);Jumlah Nilai Bruto 1 / Jumlah Nilai Nominal Obligasi Yg Diperdagangkan Di Bursa Efek / Jumlah Penghasilan Pada Form Input Yang Dibayar Sendiri;Tarif 1 / Tingkat Bunga per Tahun;PPh Yang Dipotong  1 /Jumlah PPh Pada Form Input Yang Dibayar Sendiri;Jenis Hadiah Undian 2 / Nomor Seri Obligasi ;Kode Option Tempat Penyimpanan 2;Jumlah Nilai Bruto 2 / Jumlah Harga Perolehan Bersih (tanpa Bunga) Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 2;PPh Yang Dipotong  2;Jenis Hadiah Undian 3;Kode Option Tempat Penyimpanan 3;Jumlah Nilai Bruto 3 / Jumlah Harga Penjualan Bersih (tanpa Bunga) Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 3;PPh Yang Dipotong  3;" & _
                            "Jenis Hadiah Undian 4;Kode Option Tempat Penyimpanan 4 / Kode Option Perencanaan (1) atau Pengawasan (2) atau selainnya (0) untuk BP Jasa Konstruksi poin 4;Jumlah Nilai Bruto 4 / Jumlah Diskonto Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 4;PPh Yang Dipotong  4;Jenis Hadiah Undian 5;Kode Option Tempat Penyimpanan 5 / Kode Option Perencanaan (1) atau Pengawasan (2) atau selainnya (0) untuk BP Jasa Konstruksi poin 5;Jumlah Nilai Bruto 5 / Jumlah Bunga Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 5;PPh Yang Dipotong  5;Jenis Hadiah Undian 6;Jumlah Nilai Bruto 6 / Jumlah Total Bunga atau Diskonto Obligasi Yang Diperdagangkan;Tarif 6 / Tarif PPh Final Pada Obligasi Yang Diperdagangkan Di Bursa Efek;PPh Yang Dipotong  6;Jumlah Nilai Bruto 7;Tarif 7;PPh Yang Dipotong 7;Jenis Penghasilan 8;Jumlah Nilai Bruto 8;Tarif 8;PPh Yang Dipotong 8;Jumlah PPh Yang Dipotong;Tanggal Jatuh Tempo Obligasi;Tanggal Perolehan Obligasi;Tanggal Penjualan Obligasi;" & _
                            "Holding Periode Obligasi (Hari);Time Periode Obligasi (Hari)", _
                                Me.StatusBar1)
        
    ElseIf Trim(jenisPPh) = "9" Then
    
        Call load_data_Csv("pph42_sewa", 4, 54, _
                            "NPWP KPP;Kode Form Bukti Potong / Kode Form Input PPh Yang Dibayar Sendiri;Masa Pajak;Tahun Pajak;Pembetulan;" & _
                            "NPWP WP yang Dipotong;Nama WP yang Dipotong;Alamat WP yang Dipotong;Nomor Bukti Potong / NTPN;Tanggal Bukti Potong/Tanggal SSP;Jenis Hadiah Undian 1 / Lokasi Tanah dan atau Bangunan / Nama Obligasi;Kode Option Tempat Penyimpanan 1 (Khusus F113310);Jumlah Nilai Bruto 1 / Jumlah Nilai Nominal Obligasi Yg Diperdagangkan Di Bursa Efek / Jumlah Penghasilan Pada Form Input Yang Dibayar Sendiri;Tarif 1 / Tingkat Bunga per Tahun;PPh Yang Dipotong  1 /Jumlah PPh Pada Form Input Yang Dibayar Sendiri;Jenis Hadiah Undian 2 / Nomor Seri Obligasi ;Kode Option Tempat Penyimpanan 2;Jumlah Nilai Bruto 2 / Jumlah Harga Perolehan Bersih (tanpa Bunga) Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 2;PPh Yang Dipotong  2;Jenis Hadiah Undian 3;Kode Option Tempat Penyimpanan 3;Jumlah Nilai Bruto 3 / Jumlah Harga Penjualan Bersih (tanpa Bunga) Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 3;PPh Yang Dipotong  3;" & _
                            "Jenis Hadiah Undian 4;Kode Option Tempat Penyimpanan 4 / Kode Option Perencanaan (1) atau Pengawasan (2) atau selainnya (0) untuk BP Jasa Konstruksi poin 4;Jumlah Nilai Bruto 4 / Jumlah Diskonto Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 4;PPh Yang Dipotong  4;Jenis Hadiah Undian 5;Kode Option Tempat Penyimpanan 5 / Kode Option Perencanaan (1) atau Pengawasan (2) atau selainnya (0) untuk BP Jasa Konstruksi poin 5;Jumlah Nilai Bruto 5 / Jumlah Bunga Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 5;PPh Yang Dipotong  5;Jenis Hadiah Undian 6;Jumlah Nilai Bruto 6 / Jumlah Total Bunga atau Diskonto Obligasi Yang Diperdagangkan;Tarif 6 / Tarif PPh Final Pada Obligasi Yang Diperdagangkan Di Bursa Efek;PPh Yang Dipotong  6;Jumlah Nilai Bruto 7;Tarif 7;PPh Yang Dipotong 7;Jenis Penghasilan 8;Jumlah Nilai Bruto 8;Tarif 8;PPh Yang Dipotong 8;Jumlah PPh Yang Dipotong;Tanggal Jatuh Tempo Obligasi;Tanggal Perolehan Obligasi;Tanggal Penjualan Obligasi;" & _
                            "Holding Periode Obligasi (Hari);Time Periode Obligasi (Hari)", _
                                Me.StatusBar1)
    ElseIf Trim(jenisPPh) = "10" Then
    
        Call load_data_Csv("pph42_obligasi", 4, 54, _
                            "NPWP KPP;Kode Form Bukti Potong / Kode Form Input PPh Yang Dibayar Sendiri;Masa Pajak;Tahun Pajak;Pembetulan;" & _
                            "NPWP WP yang Dipotong;Nama WP yang Dipotong;Alamat WP yang Dipotong;Nomor Bukti Potong / NTPN;Tanggal Bukti Potong/Tanggal SSP;Jenis Hadiah Undian 1 / Lokasi Tanah dan atau Bangunan / Nama Obligasi;Kode Option Tempat Penyimpanan 1 (Khusus F113310);Jumlah Nilai Bruto 1 / Jumlah Nilai Nominal Obligasi Yg Diperdagangkan Di Bursa Efek / Jumlah Penghasilan Pada Form Input Yang Dibayar Sendiri;Tarif 1 / Tingkat Bunga per Tahun;PPh Yang Dipotong  1 /Jumlah PPh Pada Form Input Yang Dibayar Sendiri;Jenis Hadiah Undian 2 / Nomor Seri Obligasi ;Kode Option Tempat Penyimpanan 2;Jumlah Nilai Bruto 2 / Jumlah Harga Perolehan Bersih (tanpa Bunga) Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 2;PPh Yang Dipotong  2;Jenis Hadiah Undian 3;Kode Option Tempat Penyimpanan 3;Jumlah Nilai Bruto 3 / Jumlah Harga Penjualan Bersih (tanpa Bunga) Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 3;PPh Yang Dipotong  3;" & _
                            "Jenis Hadiah Undian 4;Kode Option Tempat Penyimpanan 4 / Kode Option Perencanaan (1) atau Pengawasan (2) atau selainnya (0) untuk BP Jasa Konstruksi poin 4;Jumlah Nilai Bruto 4 / Jumlah Diskonto Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 4;PPh Yang Dipotong  4;Jenis Hadiah Undian 5;Kode Option Tempat Penyimpanan 5 / Kode Option Perencanaan (1) atau Pengawasan (2) atau selainnya (0) untuk BP Jasa Konstruksi poin 5;Jumlah Nilai Bruto 5 / Jumlah Bunga Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 5;PPh Yang Dipotong  5;Jenis Hadiah Undian 6;Jumlah Nilai Bruto 6 / Jumlah Total Bunga atau Diskonto Obligasi Yang Diperdagangkan;Tarif 6 / Tarif PPh Final Pada Obligasi Yang Diperdagangkan Di Bursa Efek;PPh Yang Dipotong  6;Jumlah Nilai Bruto 7;Tarif 7;PPh Yang Dipotong 7;Jenis Penghasilan 8;Jumlah Nilai Bruto 8;Tarif 8;PPh Yang Dipotong 8;Jumlah PPh Yang Dipotong;Tanggal Jatuh Tempo Obligasi;Tanggal Perolehan Obligasi;Tanggal Penjualan Obligasi;" & _
                            "Holding Periode Obligasi (Hari);Time Periode Obligasi (Hari)", _
                                Me.StatusBar1)
    ElseIf Trim(jenisPPh) = "11" Then
        Call pesan2("No report available", , vbYellow)
        'Call load_data_Csv("pph21_bwhptkp", 4, 54, _
        '                    "NPWP KPP;Kode Form Bukti Potong / Kode Form Input PPh Yang Dibayar Sendiri;Masa Pajak;Tahun Pajak;Pembetulan;" & _
        '                    "NPWP WP yang Dipotong;Nama WP yang Dipotong;Alamat WP yang Dipotong;Nomor Bukti Potong / NTPN;Tanggal Bukti Potong/Tanggal SSP;Jenis Hadiah Undian 1 / Lokasi Tanah dan atau Bangunan / Nama Obligasi;Kode Option Tempat Penyimpanan 1 (Khusus F113310);Jumlah Nilai Bruto 1 / Jumlah Nilai Nominal Obligasi Yg Diperdagangkan Di Bursa Efek / Jumlah Penghasilan Pada Form Input Yang Dibayar Sendiri;Tarif 1 / Tingkat Bunga per Tahun;PPh Yang Dipotong  1 /Jumlah PPh Pada Form Input Yang Dibayar Sendiri;Jenis Hadiah Undian 2 / Nomor Seri Obligasi ;Kode Option Tempat Penyimpanan 2;Jumlah Nilai Bruto 2 / Jumlah Harga Perolehan Bersih (tanpa Bunga) Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 2;PPh Yang Dipotong  2;Jenis Hadiah Undian 3;Kode Option Tempat Penyimpanan 3;Jumlah Nilai Bruto 3 / Jumlah Harga Penjualan Bersih (tanpa Bunga) Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 3;PPh Yang Dipotong  3;" & _
        '                    "Jenis Hadiah Undian 4;Kode Option Tempat Penyimpanan 4 / Kode Option Perencanaan (1) atau Pengawasan (2) atau selainnya (0) untuk BP Jasa Konstruksi poin 4;Jumlah Nilai Bruto 4 / Jumlah Diskonto Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 4;PPh Yang Dipotong  4;Jenis Hadiah Undian 5;Kode Option Tempat Penyimpanan 5 / Kode Option Perencanaan (1) atau Pengawasan (2) atau selainnya (0) untuk BP Jasa Konstruksi poin 5;Jumlah Nilai Bruto 5 / Jumlah Bunga Pada Obligasi Yg Diperdagangkan Di Bursa Efek;Tarif 5;PPh Yang Dipotong  5;Jenis Hadiah Undian 6;Jumlah Nilai Bruto 6 / Jumlah Total Bunga atau Diskonto Obligasi Yang Diperdagangkan;Tarif 6 / Tarif PPh Final Pada Obligasi Yang Diperdagangkan Di Bursa Efek;PPh Yang Dipotong  6;Jumlah Nilai Bruto 7;Tarif 7;PPh Yang Dipotong 7;Jenis Penghasilan 8;Jumlah Nilai Bruto 8;Tarif 8;PPh Yang Dipotong 8;Jumlah PPh Yang Dipotong;Tanggal Jatuh Tempo Obligasi;Tanggal Perolehan Obligasi;Tanggal Penjualan Obligasi;" & _
        '                    "Holding Periode Obligasi (Hari);Time Periode Obligasi (Hari)", _
        '                        Me.StatusBar1)
    ElseIf Trim(jenisPPh) = "12" Then
    
        Call load_data_Csv("pph21pesangon", 2, 16, _
                                "Masa Pajak;Tahun Pajak;Pembetulan;" & _
                                "Nomor Bukti Potong;NPWP;NIK;Nama;Alamat;Kode Pajak;" & _
                                "Jumlah Bruto;Tarif;Jumlah PPh;NPWP Pemotong;" & _
                                "Nama Pemotong;Tanggal Bukti Potong", _
                                Me.StatusBar1)
    Else
        Call pesan2("No Reports", , vbYellow)
    End If
    
    Me.Enable_Form
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
    Me.Enable_Form
End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
  '---------------------
    If dbMySQL_open = True Then
    Else
        MsgBox "Open Database Tidak Sukses", vbExclamation
    End If
    '---------------------
  
  
  'load combo
  Call load_Divisi(Me.cb_divisi, False, 1, True)
  Call load_jenisPPh(Me.cb_jenisPajak)
  Call load_KPP(Me.cb_kpp, False, 1)
  Me.opt1.Value = True
  
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
  
  Me.Width = 12540
  Me.Height = 8175
  
End Sub

Sub load_data_Csv(nmTabel1 As String, kolomAwal As Integer, kolomAkhir As Integer, header1 As String, ByRef sb1 As StatusBar)
    'k01 s/d k15
    
    
    'build rs, dengan kolom k01 s/d kxx
    'di baris pertama, inputkan nama kolom
    'di baris berikutnya, load dari tabel..
    
    Dim c As Integer
    Dim klm1 As String, sql As String
    Dim klm2
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c1 As Long
    
    
    'build rs
    klm1 = ""
    
    '---- format tambahan
    If Me.opt1.Value = True Then
    ElseIf Me.opt2.Value = True Then
        kolomAkhir = kolomAkhir + 2
        header1 = header1 & ";npwpKpp;divisi"
    End If
    '====================
    
    For c = 1 To (kolomAkhir - kolomAwal) + 1
        klm1 = klm1 & "k" & adddigit(CLng(c), 2) & ";"
    Next
    klm1 = Left(klm1, Len(klm1) - 1)
    Call create_rs2(rsGrid, klm1)
    
    '--header
    If Trim(header1) <> "" Then
        klm2 = Split(header1, ";")
        rsGrid.AddNew
        For c = 1 To (kolomAkhir - kolomAwal) + 1
            If UBound(klm2) >= c - 1 Then
                rsGrid.Fields(c - 1).Value = klm2(c - 1)
            End If
        Next
        rsGrid.Update
    End If
    
    '-- load data
    sql = create_SQL_PPH(nmTabel1, get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), Me.cb_tahun.text, _
                                Me.cb_masa.text, Me.cb_pembetulan.text)
    'sql = InputBox("", "", sql)
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    Else
        jRec = RecordCount2(rs)
        If jRec > 0 Then
            rs.MoveFirst
            c1 = 1
            Do While rs.EOF = False
                Call info(2, "Fetch csv. Run " & c1 & "/" & jRec, sb1)
                rsGrid.AddNew
                For c = kolomAwal To kolomAkhir
                    If rs.Fields.Count >= c Then
                        If nmTabel1 = "pph21tf" Then
                            If c = 11 Then
                                If cek_null(rs.Fields(c - 1)) = "N" Then
                                    'rsGrid.Fields(c - kolomAwal) = "xx1"
                                    rsGrid.Fields(c - kolomAwal) = ""
                                Else
                                    rsGrid.Fields(c - kolomAwal) = cek_null(rs.Fields(c))
                                End If
                            ElseIf c = 15 Then
                                If Left(cek_null(rs.Fields(6)), 4) = "0000" Then
                                    rsGrid.Fields(c - kolomAwal) = "Y"
                                Else
                                    rsGrid.Fields(c - kolomAwal) = "N"
                                End If
                            ElseIf c = 16 Then
                                If cek_null(rs.Fields(c)) = "" Then
                                    rsGrid.Fields(c - kolomAwal) = "5"
                                Else
                                    rsGrid.Fields(c - kolomAwal) = cek_null(rs.Fields(c))
                                End If
                            Else
                                If cek_null(rs.Fields(c)) = "" Then
                                    rsGrid.Fields(c - kolomAwal) = "0"
                                Else
                                    rsGrid.Fields(c - kolomAwal) = cek_null(rs.Fields(c))
                                End If
                            End If
                        ElseIf nmTabel1 = "pph23" Then
                            If c = 26 Then
                                If cek_null(rs.Fields(c)) = "" Then
                                    rsGrid.Fields(c - kolomAwal) = "2"
                                Else
                                    rsGrid.Fields(c - kolomAwal) = cek_null(rs.Fields(c))
                                End If
                            Else
                                If cek_null(rs.Fields(c)) = "" Then
                                    rsGrid.Fields(c - kolomAwal) = "0"
                                Else
                                    rsGrid.Fields(c - kolomAwal) = cek_null(rs.Fields(c))
                                End If
                            End If
                        ElseIf nmTabel1 = "pph42_sewa" Or nmTabel1 = "pph42_obligasi" Then
                            If c = 13 Then
                                If cek_null(rs.Fields(c)) = "" Then
                                    rsGrid.Fields(c - kolomAwal) = cek_null(rs.Fields(10))
                                Else
                                    rsGrid.Fields(c - kolomAwal) = cek_null(rs.Fields(c))
                                End If
                            Else
                                If cek_null(rs.Fields(c)) = "" Then
                                    rsGrid.Fields(c - kolomAwal) = "0"
                                Else
                                    rsGrid.Fields(c - kolomAwal) = cek_null(rs.Fields(c))
                                End If
                            End If
                        ElseIf nmTabel1 = "pph21bulanan" Then
                            If c = 10 Then
                                rsGrid.Fields(c - kolomAwal) = ""
                            Else
                                If cek_null(rs.Fields(c)) = "" Then
                                    rsGrid.Fields(c - kolomAwal) = "0"
                                Else
                                    rsGrid.Fields(c - kolomAwal) = cek_null(rs.Fields(c))
                                End If
                            End If
                        Else
                            If cek_null(rs.Fields(c)) = "" Then
                                rsGrid.Fields(c - kolomAwal) = "0"
                            Else
                                rsGrid.Fields(c - kolomAwal) = cek_null(rs.Fields(c))
                            End If
                        End If
                    End If
                    
                    '--- format2 - npwpkpp + divisi
                    If Me.opt2.Value = True Then
                        If nmTabel1 = "pph15" Then
                            rsGrid.Fields(rsGrid.Fields.Count - 2).Value = cek_null(rs.Fields(0))
                            rsGrid.Fields(rsGrid.Fields.Count - 1).Value = cek_null(rs.Fields(19))
                        ElseIf nmTabel1 = "pph21bulanan" Then
                            rsGrid.Fields(rsGrid.Fields.Count - 2).Value = cek_null(rs.Fields(0))
                            rsGrid.Fields(rsGrid.Fields.Count - 1).Value = cek_null(rs.Fields(11))
                        ElseIf nmTabel1 = "pph21tahunan" Then
                            rsGrid.Fields(rsGrid.Fields.Count - 2).Value = cek_null(rs.Fields(0))
                            rsGrid.Fields(rsGrid.Fields.Count - 1).Value = cek_null(rs.Fields(43))
                        ElseIf nmTabel1 = "pph21tf" Then
                            rsGrid.Fields(rsGrid.Fields.Count - 2).Value = cek_null(rs.Fields(0))
                            rsGrid.Fields(rsGrid.Fields.Count - 1).Value = cek_null(rs.Fields(21))
                        ElseIf nmTabel1 = "pph22" Then
                            rsGrid.Fields(rsGrid.Fields.Count - 2).Value = cek_null(rs.Fields(0))
                            rsGrid.Fields(rsGrid.Fields.Count - 1).Value = cek_null(rs.Fields(55))
                        ElseIf nmTabel1 = "pph23" Then
                            rsGrid.Fields(rsGrid.Fields.Count - 2).Value = cek_null(rs.Fields(0))
                            rsGrid.Fields(rsGrid.Fields.Count - 1).Value = cek_null(rs.Fields(81))
                        ElseIf nmTabel1 = "pph26" Then
                            rsGrid.Fields(rsGrid.Fields.Count - 2).Value = cek_null(rs.Fields(0))
                            rsGrid.Fields(rsGrid.Fields.Count - 1).Value = cek_null(rs.Fields(81))
                        ElseIf nmTabel1 = "pph42_konstruksi" Then
                            rsGrid.Fields(rsGrid.Fields.Count - 2).Value = cek_null(rs.Fields(0))
                            rsGrid.Fields(rsGrid.Fields.Count - 1).Value = cek_null(rs.Fields(55))
                        ElseIf nmTabel1 = "pph42_sewa" Or nmTabel1 = "pph42_obligasi" Then
                            rsGrid.Fields(rsGrid.Fields.Count - 2).Value = cek_null(rs.Fields(0))
                            rsGrid.Fields(rsGrid.Fields.Count - 1).Value = cek_null(rs.Fields(55))
                        End If
                    End If
                Next
                rsGrid.Update
                rs.MoveNext
                c1 = c1 + 1
            Loop
        End If
    End If
        
    '----
    
    Set Me.DataGrid1.DataSource = rsGrid
    
End Sub


Function create_SQL_PPH(nmTabel1 As String, npwp_kpp As String, kodeDivisi As String, tahunPajak As String, _
                        masaPajak As String, Pembetulan As String) As String
    
    Dim sql As String, kondisi As String

    sql = "select * from " & nmTabel1
    kondisi = ""
    
    If Trim(npwp_kpp) = "ALL" Then
    Else
        kondisi = kondisi & " NPWP_KPP = '" & Trim(npwp_kpp) & "'"
    End If
    
    If Trim(kodeDivisi) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " kode_divisi = '" & Trim(kodeDivisi) & "'"
    End If
    
    If Trim(tahunPajak) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Tahun_Pajak = '" & Trim(tahunPajak) & "'"
    End If
    
    If Trim(masaPajak) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Masa_Pajak = '" & Trim(masaPajak) & "'"
    End If
    
    If Trim(Pembetulan) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Pembetulan = '" & Trim(Pembetulan) & "'"
    End If
    
    If Trim(kondisi) = "" Then
    Else
        sql = sql & " WHERE " & kondisi
    End If
    
    
    create_SQL_PPH = sql
End Function

Private Sub Form_Resize()
    Me.Shape1.Width = Me.Width
    Me.lb_caption.Width = Me.Width
    
    If Me.Width - 645 > 0 Then Me.Frame2.Width = Me.Width - 645
    If Me.Frame2.Width - 240 > 0 Then Me.DataGrid1.Width = Me.Frame2.Width - 240
    If Me.Frame2.Width - 1695 > 0 Then Me.cmd_export.Left = Me.Frame2.Width - 1695
    
    'height
    If Me.Height - 3720 > 0 Then Me.Frame2.Height = Me.Height - 3720
    If Me.Frame2.Height - 840 > 0 Then Me.DataGrid1.Height = Me.Frame2.Height - 840
    If Me.Frame2.Height - 495 > 0 Then Me.cmd_export.Top = Me.Frame2.Height - 495
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call dbMySQL_close
End Sub

