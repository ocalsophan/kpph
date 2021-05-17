VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_impCSVssp 
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
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   6975
      Begin VB.ComboBox cb_divisi 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Text            =   "x"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Divisi"
         Enabled         =   0   'False
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   420
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
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
   Begin VB.ListBox List1 
      Height          =   1110
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Log Hasil Import File Excel. Double Klik untuk Simpan"
      Top             =   5610
      Width           =   12045
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   " 3. Isi File "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Left            =   117
      TabIndex        =   8
      Top             =   2040
      Width           =   12038
      Begin VB.CommandButton cmd_import 
         Caption         =   "Import"
         Height          =   375
         Left            =   10560
         TabIndex        =   5
         Top             =   2880
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DGrid1 
         Height          =   2460
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   11790
         _ExtentX        =   20796
         _ExtentY        =   4339
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
   End
   Begin VB.Frame Frame2 
      Caption         =   " 2. Pilih File Import "
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
      Left            =   7200
      TabIndex        =   7
      Top             =   600
      Width           =   4965
      Begin VB.TextBox txtKarakkter 
         Height          =   375
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmd_browse 
         Caption         =   "Browse"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   3405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Karakter pemisah"
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   915
         Width           =   1260
      End
   End
   Begin VB.Label lb_caption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Load CSV"
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
Attribute VB_Name = "frm_impCSVssp"
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
    'If isDataAda("mdivisi", "kodedivisi", get_kode_combo(Me.cb_divisi, "-"), cnn) = True Then
    'Else
    '    hasil = False
    '    pesan1 = pesan1 & "Divisi tidak valid"
    'End If
        
    If Trim(pesan1) = "" Then
    Else
        MsgBox pesan1
    End If
    
    cek_Isian = hasil
End Function

Private Sub cmd_browse_Click()
  Dim f As String
  Dim jmlKolom As Integer, jenisPPh As String
  
  On Error GoTo er1
  
  If cek_Isian = False Then
    Exit Sub
  End If
  
  MsgBox "Salah Pilih Format akan menampilkan hasil yang salah", vbExclamation
  Me.disable_Form
  CD.InitDir = App.Path & "\Import\"
  CD.Filter = "CSV file (*.csv;*.txt)|*.csv;*.txt"
  CD.FileName = ""
  CD.ShowOpen
  f = CD.FileName
  
  'cek inputan karakter pemisah
  If Trim(Me.txtKarakkter) = "" Then
    Call pesan2("Karakter pemisah tidak boleh kosong", , vbYellow)
    Me.Enable_Form
    Exit Sub
  End If
  '----
  
  If Trim(f) <> "" Then
    Me.Text1 = f
    If is_file_ada(f) = True Then
      'File Valid
        Call Load_Csv_2Rs(f, rs, Me.StatusBar1, Trim(Me.txtKarakkter.Text), 0)
        Me.cmd_import.Enabled = True
    Else
      MsgBox "File tidak valid", vbCritical
      Me.cmd_import.Enabled = False
    End If
  End If
  MsgBox "Jumlah data di file : " & RecordCount(rs)
  Set Me.DGrid1.DataSource = rs
  
  If RecordCount(rs) <= 0 Then
    Me.cmd_import.Enabled = False
    Me.Enable_Form
    Exit Sub
  End If
  
  jmlKolom = rs.Fields.Count
  
  Me.cmd_import.Enabled = True
    If jmlKolom = 12 Then
    Else
        Call pesan2("jumlah kolom Tidak Valid", , vbYellow)
        Me.cmd_import.Enabled = False
    End If
  
  
  '------------
  Me.Enable_Form
  Exit Sub
er1:
  MsgBox Err.Description, vbCritical
  Me.Enable_Form
End Sub

Sub disable_Form()
    Me.Frame1.Enabled = False
    Me.Frame2.Enabled = False
    Me.Frame3.Enabled = False
    Me.List1.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame1.Enabled = True
    Me.Frame2.Enabled = True
    Me.Frame3.Enabled = True
    Me.List1.Enabled = True
End Sub

Private Sub cmd_import_Click()
    Dim jRec As Long, jenisPPh As String
    Dim ps
  
    On Error GoTo er1
    
    'konfirmasi,
    ps = MsgBox("Yakin akan import Data ?" & vbCr & "Pastikan Regional Setting: Indonesia", vbYesNo)
    If ps = vbNo Then Exit Sub
  
    jRec = RecordCount(rs)
    If jRec <= 0 Then
        MsgBox "Tidak ada data"
        Exit Sub
    End If
    Me.disable_Form
    
    '---------------------
    If dbMySQL_open = True Then
    Else
        MsgBox "Open Database Tidak Sukses", vbExclamation
    End If
    '---------------------
    
    Call import_SSPpph
    
    
    Me.Enable_Form
    Exit Sub
er1:
  MsgBox Err.Description, vbCritical
  Me.Enable_Form
End Sub


Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
  Call dbMySQL_open
  
  'load combo
  'Call load_Divisi(Me.cb_divisi, , 0)
  
  'get level
  '2 : operator gedung
  '3 : UKP
  
  '---------
  
  Level1 = tbPengguna_getLevel1(frMenu1.nmLogin)
  If Level1 = 2 Then
    Me.cb_divisi.Text = tbPengguna_getDivisi(frMenu1.nmLogin)
    Me.cb_divisi.Enabled = False
  ElseIf Level1 = 3 Then
    'Me.cb_divisi.Enabled = True
  Else
    Call pesan2("Level tidak valid", , vbYellow)
    Me.cb_divisi.Enabled = False
  End If
  
  Me.Text1 = ""
  Me.txtKarakkter = ";"
  
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call dbMySQL_close
End Sub

Private Sub List1_DblClick()
  Dim pesan
  Dim namaFile As String, t1 As String
  Dim f
  Dim idx As Integer
  
  pesan = MsgBox("Simpan File Log ? ", vbYesNo)
  If pesan = vbYes Then
    Me.disable_Form
    namaFile = "d:\LogImportExcel-" & Year(Date) & "_" & Month(Date) & "_" & Day(Date) & " _ " & _
               "j" & Hour(Time) & Minute(Time) & Second(Time) & ".txt"
    Call OpenFile(namaFile, f, 2)
    For idx = 0 To List1.ListCount - 1
      List1.ListIndex = idx
      t1 = List1.Text & Chr(13) & Chr(10)
      Call writefile(f, t1)
    Next
    Call closefile(f)
    MsgBox "File export di simpan di " & namaFile, vbInformation
    Me.Enable_Form
  End If
End Sub


Sub import_SSPpph()
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long
  
    Dim npwp_kpp As String, Kode_Form As String, Masa_Pajak_SSP As String
    Dim Tahun_Pajak_SSP As String, Pembetulan As String, NTPN As String
    Dim Tanggal_Setor_SSP As Date, Jumlah_SSP As Currency, Kode_KAP As String
    Dim Kode_Jenis_Setoran As String, Jenis_Pajak As String, kode_divisi As String
                    
    Dim return1 As Integer
    '--------------------------
    Dim t As String, ps, sql As String
    Dim data_Valid As Boolean
  
    jRec = RecordCount(rs)
    If jRec <= 0 Then
        MsgBox "Tidak ada data"
        Exit Sub
    End If
      
    rs.MoveFirst
    c = 0
    jml_Insert = 0
    jml_Update = 0
    Me.List1.Clear
    Do While rs.EOF = False
        c = c + 1
        Call info(1, "Run Import " & c & " of " & jRec & ". Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update, Me.StatusBar1)
        data_Valid = True
    
        'ssssssssssssssssss
        npwp_kpp = cek_null(rs(0))
        Jenis_Pajak = cek_null(rs(1))
        kode_divisi = cek_null(rs(2))
        
        Kode_Form = cek_null(rs(3))
        If (Kode_Form = "F113204") Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Kode Form tidak valid")
        End If
        
        
        Masa_Pajak_SSP = cek_null(rs(4))
        Tahun_Pajak_SSP = cek_null(rs(5))
        Pembetulan = cek_null(rs(6))
        NTPN = cek_null(rs(7))
        
        Tanggal_Setor_SSP = cek_Date(rs(8))
        
        '-----------
        'If CStr(Year(Tanggal_Setor_SSP)) = Tahun_Pajak_SSP Then
        'Else
        '    data_Valid = False
        '    Call setListInfo(Me.List1, "Data ke " & c & " Tahun Pajak tidak valid")
        'End If
        '
        'If adddigit(CLng(Month(Tanggal_Setor_SSP)), 2) = adddigit(cek_Lng(Masa_Pajak_SSP), 2) Then
        'Else
        '     data_Valid = False
        '    Call setListInfo(Me.List1, "Data ke " & c & " Masa Pajak tidak valid")
        'End If
        '-----------
        
        
        Jumlah_SSP = cek_Money(rs(9))
        Kode_KAP = cek_null(rs(10))
        Kode_Jenis_Setoran = cek_null(rs(11))
        '------------------
     
        If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " NPWP_KPP tidak valid"
        ElseIf Trim(NTPN) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " NTPN tidak valid"
        ElseIf Jumlah_SSP = 0 Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nilai SSP 0"
        ElseIf Trim(kode_divisi) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Kode Divisi tidak valid"
        ElseIf Trim(Jenis_Pajak) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Jenis Pajak tidak valid"
        End If
    
        If data_Valid = True Then
            return1 = tbSSPpph_insert(npwp_kpp, Kode_Form, Masa_Pajak_SSP, Tahun_Pajak_SSP, Pembetulan, NTPN, _
                                    Tanggal_Setor_SSP, Jumlah_SSP, Kode_KAP, Kode_Jenis_Setoran, _
                                    Jenis_Pajak, kode_divisi)
            If return1 = 1 Then
                jml_Insert = jml_Insert + 1
            ElseIf return1 = 2 Then
                jml_Update = jml_Update + 1
            Else
                Call pesan2("Insert error", , vbYellow)
                Exit Sub
            End If
        End If
        rs.MoveNext
    Loop
    MsgBox "Proses Import Selesai. Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update, vbInformation

End Sub

