VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_impCSV 
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
      TabIndex        =   13
      Top             =   600
      Width           =   6975
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
         TabIndex        =   15
         Top             =   780
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Divisi"
         Height          =   210
         Left            =   240
         TabIndex        =   14
         Top             =   420
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   12
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
      TabIndex        =   7
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
      TabIndex        =   9
      Top             =   2040
      Width           =   12038
      Begin VB.CommandButton cmd_import 
         Caption         =   "Import"
         Height          =   375
         Left            =   10560
         TabIndex        =   6
         Top             =   2880
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DGrid1 
         Height          =   2460
         Left            =   120
         TabIndex        =   10
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
      TabIndex        =   8
      Top             =   600
      Width           =   4965
      Begin VB.CommandButton cmd_template 
         BackColor       =   &H00FFFFC0&
         Caption         =   "- File Template"
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "file template import"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtKarakkter 
         Height          =   375
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmd_browse 
         Caption         =   "Browse"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   360
         Width           =   3405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Karakter pemisah"
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   915
         Width           =   1260
      End
   End
   Begin VB.Label lb_caption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Load CSV - Data SPT"
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
Attribute VB_Name = "frm_impCSV"
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
    If t = "1" Or t = "2" Or t = "3" Or t = "4" Or t = "5" Or t = "6" _
        Or t = "7" Or t = "8" Or t = "9" Or t = "10" Or t = "11" Or t = "12" Then
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

Private Sub cmd_browse_Click()
  Dim f As String
  Dim jmlKolom As Integer, jenisPPh As String
  
  On Error GoTo er1
  
  '---------------------
    If dbMySQL_open = True Then
    Else
        MsgBox "Open Database Tidak Sukses", vbExclamation
    End If
    '---------------------
  
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
        Call Load_Csv_2Rs(f, rs, Me.StatusBar1, Trim(Me.txtKarakkter.text), 0)
        Me.cmd_import.Enabled = True
    Else
      MsgBox "File tidak valid", vbCritical
      Me.cmd_import.Enabled = False
    End If
  End If
  MsgBox "Jumlah data di file : " & RecordCount(rs)
  Set Me.DGrid1.DataSource = rs
  
  'jumlah kolom
  '1 : 15 kolom
  '2 : 51 kolom
  '3 : 19
  '4 : 9
  '5 : 51
  '6 : 77
  '7 : 51
  
  
  If RecordCount(rs) <= 0 Then
    Me.cmd_import.Enabled = False
    Me.Enable_Form
    Exit Sub
  End If
  
  jenisPPh = get_kode_combo(Me.cb_jenisPajak, ".")
  jmlKolom = rs.Fields.Count
  
  Me.cmd_import.Enabled = True
  If Trim(jenisPPh) = "1" Then
    If jmlKolom = 20 Then
    Else
        Call pesan2("jumlah kolom Tidak Valid", , vbYellow)
        Me.cmd_import.Enabled = False
    End If
  ElseIf Trim(jenisPPh) = "2" Then
    If jmlKolom = 82 Then
    Else
        Call pesan2("jumlah kolom Tidak Valid", , vbYellow)
        Me.cmd_import.Enabled = False
    End If
  ElseIf Trim(jenisPPh) = "3" Then
    If jmlKolom = 22 Then
    Else
        Call pesan2("jumlah kolom Tidak Valid", , vbYellow)
        Me.cmd_import.Enabled = False
    End If
  ElseIf Trim(jenisPPh) = "4" Then
    If jmlKolom = 12 Then
    Else
        Call pesan2("jumlah kolom Tidak Valid", , vbYellow)
        Me.cmd_import.Enabled = False
    End If
  ElseIf Trim(jenisPPh) = "5" Then
    If jmlKolom = 44 Then
    Else
        Call pesan2("jumlah kolom Tidak Valid", , vbYellow)
        Me.cmd_import.Enabled = False
    End If
  ElseIf Trim(jenisPPh) = "6" Then
    If jmlKolom = 55 Then
    Else
        Call pesan2("jumlah kolom Tidak Valid", , vbYellow)
        Me.cmd_import.Enabled = False
    End If
  ElseIf Trim(jenisPPh) = "7" Then
    If jmlKolom = 82 Then
    Else
        Call pesan2("jumlah kolom Tidak Valid", , vbYellow)
        Me.cmd_import.Enabled = False
    End If
  ElseIf Trim(jenisPPh) = "8" Then
    If jmlKolom = 56 Then
    Else
        Call pesan2("jumlah kolom Tidak Valid", , vbYellow)
        Me.cmd_import.Enabled = False
    End If
  ElseIf Trim(jenisPPh) = "9" Or Trim(jenisPPh) = "10" Then
    If jmlKolom = 56 Then
    Else
        Call pesan2("jumlah kolom Tidak Valid", , vbYellow)
        Me.cmd_import.Enabled = False
    End If
  ElseIf Trim(jenisPPh) = "11" Then
    If jmlKolom = 7 Then
    Else
        Call pesan2("jumlah kolom Tidak Valid", , vbYellow)
        Me.cmd_import.Enabled = False
    End If
  ElseIf Trim(jenisPPh) = "12" Then
    If jmlKolom = 18 Then
    Else
        Call pesan2("jumlah kolom Tidak Valid", , vbYellow)
        Me.cmd_import.Enabled = False
    End If
  Else
    Call pesan2("jumlah kolom Tidak Valid", , vbYellow)
    Me.cmd_import.Enabled = False
  End If
  
  
  '------------
  Me.Enable_Form
  Exit Sub
er1:
  MsgBox Err.DESCRIPTION, vbCritical
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
    
    jenisPPh = get_kode_combo(Me.cb_jenisPajak, ".")
    If Trim(jenisPPh) = "1" Then
        Call import_pph15
    ElseIf Trim(jenisPPh) = "2" Then
        import_pph23
    ElseIf Trim(jenisPPh) = "3" Then
        import_pph21tF
    ElseIf Trim(jenisPPh) = "4" Then
        import_pph21Bulanan
    ElseIf Trim(jenisPPh) = "5" Then
        import_pph21Tahunan
    ElseIf Trim(jenisPPh) = "6" Then
        import_pph22
    ElseIf Trim(jenisPPh) = "7" Then
        import_pph26
    ElseIf Trim(jenisPPh) = "8" Then
        import_pph42_Konstruksi
    ElseIf Trim(jenisPPh) = "9" Then
        import_pph42_Sewa
    ElseIf Trim(jenisPPh) = "10" Then
        import_pph42_obligasi
    ElseIf Trim(jenisPPh) = "11" Then
        import_pph21_bwhptkp
    ElseIf Trim(jenisPPh) = "12" Then
        import_pph21pesangon
    Else
        Call pesan2("nothing to do..", , vbYellow)
    End If
    
    
    
    Me.Enable_Form
    Exit Sub
er1:
  MsgBox Err.DESCRIPTION, vbCritical
  Me.Enable_Form
End Sub


Sub import_pph15()
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long
  
    Dim Kode_Form As String, Masa_Pajak As String, Tahun_Pajak As String, Pembetulan As String
    Dim npwp_wp As String, Nama_WP As String, Alamat_WP As String, Nomor_Bukti_Potong As String
    Dim Tanggal_Bukti_Potong As Date, negara_sumber_penghasilan As String
    Dim kode_option_penghasilan As String, Jumlah_Bruto As Currency, Tarif As String
    Dim pph_dipotong As Currency, invoice_ket As String, kode_divisi As String, npwp_kpp As String
    Dim kd_proyek As String, nott As String, nofaktur As String, email As String
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
        npwp_kpp = cleanNpwp(rs(0))
        
        kd_proyek = cleanNpwp(rs(1))
        nott = cleanNpwp(rs(2))
        nofaktur = cleanNpwp(rs(3))
        
        
        Kode_Form = cleanStr(rs(4))
        If Kode_Form = "F113314" Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Kode Form tidak valid")
        End If
        Masa_Pajak = cleanStr(rs(5))
        Tahun_Pajak = cleanStr(rs(6))
        Pembetulan = cleanStr(rs(7))
        npwp_wp = cleanStr(rs(8))
        If checkNPWP(npwp_wp) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP tidak valid")
        End If
        
        Nama_WP = cleanStr(rs(9))
        Alamat_WP = cleanStr(rs(10))
        
        Call cek_npwpWP(npwp_wp, Nama_WP, Alamat_WP)
        
        Nomor_Bukti_Potong = cleanStr(rs(11))
        Tanggal_Bukti_Potong = cek_Date(rs(12))
        If CStr(Year(Tanggal_Bukti_Potong)) = Tahun_Pajak Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Tahun Pajak tidak valid")
        End If
        
        If adddigit(CLng(Month(Tanggal_Bukti_Potong)), 2) = adddigit(cek_Lng(Masa_Pajak), 2) Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Masa Pajak tidak valid")
        End If
        
        negara_sumber_penghasilan = cleanStr(rs(13))
        kode_option_penghasilan = cleanStr(rs(14))
        Jumlah_Bruto = cek_Money(rs(15))
        Tarif = cek_null(rs(16))
        pph_dipotong = cek_Money(rs(17))
        invoice_ket = cleanStr(rs(18))
        email = cleanStr(rs(19))
        
        kode_divisi = get_kode_combo(Me.cb_divisi, "-")
        '------------------
     
        If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP_KPP tidak valid")
        ElseIf Trim(Nomor_Bukti_Potong) = "" Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Nomor bukti potong tidak valid")
        ElseIf pph_dipotong = 0 Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Nilai PPH 0")
        End If
    
        If data_Valid = True Then
        
            return1 = tbPph15_insert(Kode_Form, Masa_Pajak, Tahun_Pajak, Pembetulan, npwp_wp, Nama_WP, _
                                Alamat_WP, Nomor_Bukti_Potong, Tanggal_Bukti_Potong, _
                                negara_sumber_penghasilan, kode_option_penghasilan, Jumlah_Bruto, _
                                Tarif, pph_dipotong, invoice_ket, kode_divisi, npwp_kpp, kd_proyek, nott, nofaktur, email)
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

Sub import_pph23()
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long
  
    Dim npwp_kpp As String, Kode_Form As String, Masa_Pajak As String, Tahun_Pajak As String
    Dim Pembetulan As String, npwp_wp As String, Nama_WP As String, Alamat_WP As String
    Dim Nomor_Bukti_Potong As String, Tanggal_Bukti_Potong As Date
    Dim Nilai_Bruto_1 As Currency, Tarif_1 As String, PPh_Yang_Dipotong__1 As Currency
    Dim Nilai_Bruto_2 As Currency, Tarif_2 As String, PPh_Yang_Dipotong__2 As Currency
    Dim Nilai_Bruto_3 As Currency, Tarif_3 As String, PPh_Yang_Dipotong__3 As Currency
    Dim Nilai_Bruto_4 As Currency, Tarif_4 As String, PPh_Yang_Dipotong__4 As Currency
    Dim Nilai_Bruto_5 As Currency, Tarif_5 As String, PPh_Yang_Dipotong__5 As Currency
    Dim Nilai_Bruto_6a As Currency, Tarif_6a As String, PPh_Yang_Dipotong__6a As Currency
    Dim Nilai_Bruto_6b As Currency, Tarif_6b As String, PPh_Yang_Dipotong__6b As Currency
    Dim Nilai_Bruto_6c As Currency, Tarif_6c As String, PPh_Yang_Dipotong__6c As Currency
    Dim Kode_Jasa_6d1 As String, Nilai_Bruto_6d1 As Currency, Tarif_6d1 As String, PPh_Yang_Dipotong__6d1 As Currency
    Dim Jumlah_Nilai_Bruto_ As Currency, Jumlah_PPh_Yang_Dipotong As Currency, kode_divisi As String
    Dim kd_proyek As String, nott As String, nofaktur As String, email As String
                    
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
        npwp_kpp = cleanNpwp(rs(0))
        
        kd_proyek = cleanNpwp(rs(1))
        nott = cleanNpwp(rs(2))
        nofaktur = cleanNpwp(rs(3))
        
        Kode_Form = cleanStr(rs(4))
        If Kode_Form = "F113306" Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Kode Form tidak valid")
        End If
        Masa_Pajak = cleanStr(rs(5))
        Tahun_Pajak = cleanStr(rs(6))
        Pembetulan = cleanStr(rs(7))
        npwp_wp = cleanStr(rs(8))
        If checkNPWP(npwp_wp) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP tidak valid")
        End If
        
        Nama_WP = cleanStr(rs(9))
        Alamat_WP = cleanStr(rs(10))
        
        Call cek_npwpWP(npwp_wp, Nama_WP, Alamat_WP)
        
        Nomor_Bukti_Potong = cleanStr(rs(11))
        Tanggal_Bukti_Potong = cek_Date(rs(12))
        
        
        '-----------
        If CStr(Year(Tanggal_Bukti_Potong)) = Tahun_Pajak Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Tahun Pajak tidak valid")
        End If
        
        If adddigit(CLng(Month(Tanggal_Bukti_Potong)), 2) = adddigit(cek_Lng(Masa_Pajak), 2) Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Masa Pajak tidak valid")
        End If
        '-----------
        
        If CStr(Year(Tanggal_Bukti_Potong)) = Tahun_Pajak Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Tahun Pajak tidak valid")
        End If
        
        If adddigit(CLng(Month(Tanggal_Bukti_Potong)), 2) = adddigit(cek_Lng(Masa_Pajak), 2) Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Masa Pajak tidak valid")
        End If
        
        Nilai_Bruto_1 = cek_Money(rs(13))
        Tarif_1 = cleanStr(rs(14))
        PPh_Yang_Dipotong__1 = cek_Money(rs(15))
        
        Nilai_Bruto_2 = cek_Money(rs(16))
        Tarif_2 = cleanStr(rs(17))
        PPh_Yang_Dipotong__2 = cek_Money(rs(18))
        
        Nilai_Bruto_3 = cek_Money(rs(19))
        Tarif_3 = cleanStr(rs(20))
        PPh_Yang_Dipotong__3 = cek_Money(rs(21))
        
        Nilai_Bruto_4 = cek_Money(rs(22))
        Tarif_4 = cleanStr(rs(23))
        PPh_Yang_Dipotong__4 = cek_Money(rs(24))
        
        Nilai_Bruto_5 = cek_Money(rs(25))
        Tarif_5 = cleanStr(rs(26))
        PPh_Yang_Dipotong__5 = cek_Money(rs(27))

        Nilai_Bruto_6a = cek_Money(rs(28))
        Tarif_6a = cleanStr(rs(29))
        PPh_Yang_Dipotong__6a = cek_Money(rs(30))
        
        Nilai_Bruto_6b = cek_Money(rs(31))
        Tarif_6b = cleanStr(rs(32))
        PPh_Yang_Dipotong__6b = cek_Money(rs(33))
        
        Nilai_Bruto_6c = cek_Money(rs(34))
        Tarif_6c = cleanStr(rs(35))
        PPh_Yang_Dipotong__6c = cek_Money(rs(36))
        
        
        Kode_Jasa_6d1 = cleanStr(rs(55))
        Nilai_Bruto_6d1 = cek_Money(rs(56))
        Tarif_6d1 = cleanStr(rs(57))
        PPh_Yang_Dipotong__6d1 = cek_Money(rs(58))
        
        Jumlah_Nilai_Bruto_ = cek_Money(rs(79))
        Jumlah_PPh_Yang_Dipotong = cek_Money(rs(80))
        email = cleanStr(rs(81))
        
        kode_divisi = get_kode_combo(Me.cb_divisi, "-")
        '------------------
     
        If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " NPWP_KPP tidak valid"
        ElseIf Trim(Nomor_Bukti_Potong) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nomor bukti potong tidak valid"
        ElseIf Jumlah_PPh_Yang_Dipotong = 0 Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nilai PPH 0"
        End If
    
        If data_Valid = True Then
        
            return1 = tbPph23_insert(npwp_kpp, Kode_Form, Masa_Pajak, Tahun_Pajak, Pembetulan, npwp_wp, Nama_WP, _
                                Alamat_WP, Nomor_Bukti_Potong, Tanggal_Bukti_Potong, _
                                Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1, _
                                Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, _
                                Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, _
                                Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, _
                                Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, _
                                Nilai_Bruto_6a, Tarif_6a, PPh_Yang_Dipotong__6a, _
                                Nilai_Bruto_6b, Tarif_6b, PPh_Yang_Dipotong__6b, _
                                Nilai_Bruto_6c, Tarif_6c, PPh_Yang_Dipotong__6c, _
                                Kode_Jasa_6d1, Nilai_Bruto_6d1, Tarif_6d1, PPh_Yang_Dipotong__6d1, _
                                Jumlah_Nilai_Bruto_, Jumlah_PPh_Yang_Dipotong, kode_divisi, kd_proyek, nott, nofaktur, email)
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

Sub import_pph21tF()
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long, jml_Skip As Long
  
    Dim npwp_kpp As String, Masa_Pajak As String, Tahun_Pajak As String
    Dim Pembetulan As String, Nomor_Bukti_Potong As String, npwp As String
    Dim NIK As String, nama As String, alamat As String, WP_Luar_Negeri As String
    Dim Kode_Negara As String, Kode_Pajak As String, Jumlah_Bruto As Currency
    Dim Jumlah_DPP As Currency, Tanpa_NPWP As Currency, Tarif As String
    Dim Jumlah_PPh As Currency, NPWP_Pemotong As String, Nama_Pemotong As String
    Dim Tanggal_Bukti_Potong As Date, kode_divisi As String
    Dim kd_proyek As String, email As String
                    
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
    jml_Skip = 0
    
    Me.List1.Clear
    Do While rs.EOF = False
        c = c + 1
        Call info(1, "Run Import " & c & " of " & jRec & ". Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update, Me.StatusBar1)
        data_Valid = True
    
        'ssssssssssssssssss
        npwp_kpp = cleanNpwp(rs(0))
        kd_proyek = cleanNpwp(rs(1))
        
        Masa_Pajak = cek_null(rs(2))
        Tahun_Pajak = cek_null(rs(3))
        Pembetulan = cek_null(rs(4))
        Nomor_Bukti_Potong = cek_null(rs(5))
        npwp = cek_null(rs(6))
        
        If checkNPWP(npwp) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP tidak valid")
        End If
        
        NIK = cek_null(rs(7))
        nama = cek_null(rs(8))
        alamat = cek_null(rs(9))
        
        Call cek_npwpWP(npwp, nama, alamat)
        
        WP_Luar_Negeri = cek_null(rs(10))
        Kode_Negara = cek_null(rs(11))
        
        Kode_Pajak = cek_null(rs(12))
        Jumlah_Bruto = cek_Money(rs(13))
        Jumlah_DPP = cek_Money(rs(14))
        Tanpa_NPWP = cek_Money(rs(15))
        Tarif = cek_null(rs(16))
        
        Jumlah_PPh = cek_Money(rs(17))
        NPWP_Pemotong = cek_null(rs(18))
        Nama_Pemotong = cek_null(rs(19))
        Tanggal_Bukti_Potong = cek_Date(rs(20))
        email = cleanStr(rs(21))
        
        If CStr(Year(Tanggal_Bukti_Potong)) = Tahun_Pajak Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Tahun Pajak tidak valid")
        End If
        
        If adddigit(CLng(Month(Tanggal_Bukti_Potong)), 2) = adddigit(cek_Lng(Masa_Pajak), 2) Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Masa Pajak tidak valid")
        End If
        
        
        kode_divisi = get_kode_combo(Me.cb_divisi, "-")
        '------------------
     
        If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " NPWP_KPP tidak valid"
        ElseIf Trim(Nomor_Bukti_Potong) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nomor bukti potong tidak valid"
        ElseIf Jumlah_DPP = 0 Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nilai DPP 0"
        End If
    
        If data_Valid = True Then
        
            return1 = tbPph21tf_insert(npwp_kpp, Masa_Pajak, Tahun_Pajak, Pembetulan, Nomor_Bukti_Potong, _
                                npwp, NIK, nama, alamat, WP_Luar_Negeri, Kode_Negara, Kode_Pajak, _
                                Jumlah_Bruto, Jumlah_DPP, Tanpa_NPWP, Tarif, Jumlah_PPh, NPWP_Pemotong, _
                                Nama_Pemotong, Tanggal_Bukti_Potong, kode_divisi, kd_proyek, email)
            If return1 = 1 Then
                jml_Insert = jml_Insert + 1
            ElseIf return1 = 2 Then
                jml_Update = jml_Update + 1
            ElseIf return1 = 3 Then
                jml_Skip = jml_Skip + 1
            Else
                Call pesan2("Insert error", , vbYellow)
                Exit Sub
            End If
        End If
        rs.MoveNext
    Loop
    MsgBox "Proses Import Selesai. Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update & " Jml Skip: " & jml_Skip, vbInformation

End Sub

Sub import_pph21pesangon()
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long, jml_Skip As Long
  
    Dim npwp_kpp As String, Masa_Pajak As String, Tahun_Pajak As String
    Dim Pembetulan As String, Nomor_Bukti_Potong As String, npwp As String
    Dim NIK As String, nama As String, alamat As String
    Dim Kode_Pajak As String, Jumlah_Bruto As Currency
    Dim Tarif As String
    Dim Jumlah_PPh As Currency, NPWP_Pemotong As String, Nama_Pemotong As String
    Dim Tanggal_Bukti_Potong As Date, kode_divisi As String
    Dim kd_proyek As String, email As String
                    
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
    jml_Skip = 0
    
    Me.List1.Clear
    Do While rs.EOF = False
        c = c + 1
        Call info(1, "Run Import " & c & " of " & jRec & ". Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update, Me.StatusBar1)
        data_Valid = True
    
        'ssssssssssssssssss
        npwp_kpp = cleanNpwp(rs(0))
        kd_proyek = cleanNpwp(rs(1))
        
        Masa_Pajak = cek_null(rs(2))
        Tahun_Pajak = cek_null(rs(3))
        Pembetulan = cek_null(rs(4))
        Nomor_Bukti_Potong = cek_null(rs(5))
        npwp = cek_null(rs(6))
        
        If checkNPWP(npwp) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP tidak valid")
        End If
        
        NIK = cek_null(rs(7))
        nama = cek_null(rs(8))
        alamat = cek_null(rs(9))
        
        Call cek_npwpWP(npwp, nama, alamat)
        
        
        Kode_Pajak = cek_null(rs(10))
        Jumlah_Bruto = cek_Money(rs(11))
        Tarif = cek_null(rs(12))
        
        Jumlah_PPh = cek_Money(rs(13))
        NPWP_Pemotong = cek_null(rs(14))
        Nama_Pemotong = cek_null(rs(15))
        Tanggal_Bukti_Potong = cek_Date(rs(16))
        email = cleanStr(rs(17))
        
        If CStr(Year(Tanggal_Bukti_Potong)) = Tahun_Pajak Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Tahun Pajak tidak valid")
        End If
        
        If adddigit(CLng(Month(Tanggal_Bukti_Potong)), 2) = adddigit(cek_Lng(Masa_Pajak), 2) Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Masa Pajak tidak valid")
        End If
        
        
        kode_divisi = get_kode_combo(Me.cb_divisi, "-")
        '------------------
     
        If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " NPWP_KPP tidak valid"
        ElseIf Trim(Nomor_Bukti_Potong) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nomor bukti potong tidak valid"
        ElseIf Jumlah_Bruto = 0 Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nilai Bruto 0"
        End If
    
        If data_Valid = True Then
        
            return1 = tbPph21pesangon_insert(npwp_kpp, Masa_Pajak, Tahun_Pajak, Pembetulan, Nomor_Bukti_Potong, _
                                npwp, NIK, nama, alamat, Kode_Pajak, _
                                Jumlah_Bruto, Tarif, Jumlah_PPh, NPWP_Pemotong, _
                                Nama_Pemotong, Tanggal_Bukti_Potong, kode_divisi, kd_proyek, email)
            If return1 = 1 Then
                jml_Insert = jml_Insert + 1
            ElseIf return1 = 2 Then
                jml_Update = jml_Update + 1
            ElseIf return1 = 3 Then
                jml_Skip = jml_Skip + 1
            Else
                Call pesan2("Insert error", , vbYellow)
                Exit Sub
            End If
        End If
        rs.MoveNext
    Loop
    MsgBox "Proses Import Selesai. Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update & " Jml Skip: " & jml_Skip, vbInformation

End Sub



Private Sub cmd_template_Click()
    Dim myPath As String
    myPath = App.Path & "\temp"
    Call Shell("explorer " & myPath, vbNormalFocus)
    'Call pesan2("Silahkan Cek window/jendela yang terbuka", 1000)
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
  Call load_Divisi(Me.cb_divisi, , 0)
  Call load_jenisPPh(Me.cb_jenisPajak)
  
  
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
      t1 = List1.text & Chr(13) & Chr(10)
      Call writefile(f, t1)
    Next
    Call closefile(f)
    MsgBox "File export di simpan di " & namaFile, vbInformation
    Me.Enable_Form
  End If
End Sub

Sub import_pph21Bulanan()
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long, jml_Skip As Long
  
    Dim npwp_kpp As String, Masa_Pajak As String, Tahun_Pajak As String
    Dim Pembetulan As String, npwp As String, nama As String, Kode_Pajak As String
    Dim Jumlah_Bruto As Currency, Jumlah_PPh As Currency, Kode_Negara As String
    Dim kode_divisi As String
    Dim kd_proyek As String, email As String
                    
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
    jml_Skip = 0
    
    Me.List1.Clear
    Do While rs.EOF = False
        c = c + 1
        Call info(1, "Run Import " & c & " of " & jRec & ". Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update, Me.StatusBar1)
        data_Valid = True
    
        'ssssssssssssssssss
        npwp_kpp = cek_null(rs(0))
        kd_proyek = cleanNpwp(rs(1))
        
        Masa_Pajak = cek_null(rs(2))
        Tahun_Pajak = cek_null(rs(3))
        Pembetulan = cek_null(rs(4))
        npwp = cek_null(rs(5))
        
        If checkNPWP(npwp) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP tidak valid")
        End If
        
        nama = cleanStr(cek_null(rs(6)))
        
        Call cek_npwpWP(npwp, nama, "")
        
        Kode_Pajak = cek_null(rs(7))
        Jumlah_Bruto = cek_Money(rs(8))
        Jumlah_PPh = cek_Money(rs(9))
        Kode_Negara = cek_null(rs(10))
        email = cleanStr(rs(11))
        
        kode_divisi = get_kode_combo(Me.cb_divisi, "-")
        '------------------
     
        If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP_KPP tidak valid")
        ElseIf Trim(npwp) & Trim(nama) = "" Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP/NAMA WP tidak valid")
        ElseIf Jumlah_PPh = 0 Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & "  Nilai PPH 0")
        ElseIf tbMkaryawan_isDataAda2(npwp, nama) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & "  NPWP / NAMA tidak terdaftar di master karyawan")
        End If
    
        If data_Valid = True Then
        
            return1 = tbPph21Bulanan_insert(npwp_kpp, Masa_Pajak, Tahun_Pajak, Pembetulan, npwp, _
                                nama, Kode_Pajak, Jumlah_Bruto, Jumlah_PPh, Kode_Negara, _
                                kode_divisi, kd_proyek, email)
            If return1 = 1 Then
                jml_Insert = jml_Insert + 1
            ElseIf return1 = 2 Then
                jml_Update = jml_Update + 1
            ElseIf return1 = 3 Then
                jml_Skip = jml_Skip + 1
                Call setListInfo(Me.List1, "Data ke " & c & "  sudah ada")
            Else
                Call pesan2("Insert error", , vbYellow)
                Exit Sub
            End If
        End If
        rs.MoveNext
    Loop
    MsgBox "Proses Import Selesai. Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update & " Jml Skip: " & jml_Skip, vbInformation

End Sub

Sub import_pph21_bwhptkp()
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long, jml_Skip As Long
  
    Dim npwp_kpp As String, Masa_Pajak As String, Tahun_Pajak As String
    Dim Pembetulan As String, Jumlah_karyawan As Integer
    Dim Jumlah_Bruto As Currency
    Dim kode_divisi As String
    Dim kd_proyek As String, email As String
                    
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
    jml_Skip = 0
    
    Me.List1.Clear
    Do While rs.EOF = False
        c = c + 1
        Call info(1, "Run Import " & c & " of " & jRec & ". Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update, Me.StatusBar1)
        data_Valid = True
    
        'ssssssssssssssssss
        npwp_kpp = cek_null(rs(0))
        kd_proyek = cleanNpwp(rs(1))
        
        Masa_Pajak = cek_null(rs(2))
        Tahun_Pajak = cek_null(rs(3))
        Pembetulan = cek_null(rs(4))
        
        Jumlah_karyawan = cek_Int(cek_null(rs(5)))
        
        Jumlah_Bruto = cek_Money(rs(6))
        
        kode_divisi = get_kode_combo(Me.cb_divisi, "-")
        '------------------
     
        If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP_KPP tidak valid")
        ElseIf Jumlah_karyawan <= 0 Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Jumlah_karyawan tidak valid")
        ElseIf Jumlah_Bruto = 0 Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & "  Jumlah_Bruto 0")
        End If
    
        If data_Valid = True Then
        
            return1 = tbPph21bwhptkp_insert(npwp_kpp, Masa_Pajak, Tahun_Pajak, Pembetulan, _
                                Jumlah_karyawan, Jumlah_Bruto, kode_divisi, kd_proyek)
            If return1 = 1 Then
                jml_Insert = jml_Insert + 1
            ElseIf return1 = 2 Then
                jml_Update = jml_Update + 1
            ElseIf return1 = 3 Then
                jml_Skip = jml_Skip + 1
                Call setListInfo(Me.List1, "Data ke " & c & "  sudah ada")
            Else
                Call pesan2("Insert error", , vbYellow)
                Exit Sub
            End If
        End If
        rs.MoveNext
    Loop
    MsgBox "Proses Import Selesai. Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update & " Jml Skip: " & jml_Skip, vbInformation

End Sub


Sub import_pph21Tahunan()
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long, jml_Skip As Long
  
    Dim npwp_kpp As String, Masa_Pajak As String, Tahun_Pajak As String
    Dim Pembetulan As String, Nomor_Bukti_Potong As String, Masa_Perolehan_Awal As String
    Dim Masa_Perolehan_Akhir As String, npwp As String, NIK As String, nama As String
    Dim alamat As String, jenis_kelamin As String, Status_PTKP As String
    Dim Jumlah_Tanggungan As String, Nama_Jabatan As String, WP_Luar_Negeri As String
    Dim Kode_Negara As String, Kode_Pajak As String, Jumlah_1 As Currency
    Dim Jumlah_2 As Currency, Jumlah_3 As Currency, Jumlah_4 As Currency
    Dim Jumlah_5 As Currency, Jumlah_6 As Currency, Jumlah_7 As Currency
    Dim Jumlah_8 As Currency, Jumlah_9 As Currency, Jumlah_10 As Currency
    Dim Jumlah_11 As Currency, Jumlah_12 As Currency, Jumlah_13 As Currency
    Dim Jumlah_14 As Currency, Jumlah_15 As Currency, Jumlah_16 As Currency
    Dim Jumlah_17 As Currency, Jumlah_18 As Currency, Jumlah_19 As Currency
    Dim Jumlah_20 As Currency, Status_Pindah As String, NPWP_Pemotong As String
    Dim Nama_Pemotong As String, Tanggal_Bukti_Potong As Date, kode_divisi As String
    Dim kd_proyek As String, email As String
                    
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
    jml_Skip = 0
    
    Me.List1.Clear
    Do While rs.EOF = False
        c = c + 1
        Call info(1, "Run Import " & c & " of " & jRec & ". Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update, Me.StatusBar1)
        data_Valid = True
    
        'ssssssssssssssssss
        npwp_kpp = cleanNpwp(rs(0))
        kd_proyek = cleanNpwp(rs(1))
        
        Masa_Pajak = cek_null(rs(2))
        Tahun_Pajak = cek_null(rs(3))
        Pembetulan = cek_null(rs(4))
        Nomor_Bukti_Potong = cek_null(rs(5))
        Masa_Perolehan_Awal = cek_null(rs(6))
        Masa_Perolehan_Akhir = cek_null(rs(7))
        npwp = cek_null(rs(8))
        
        If checkNPWP(npwp) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP tidak valid")
        End If
        
        NIK = cek_null(rs(9))
        nama = cleanStr(cek_null(rs(10)))
        alamat = cleanStr(cek_null(rs(11)))
        
        Call cek_npwpWP(npwp, nama, alamat)
        
        jenis_kelamin = cek_null(rs(12))
        Status_PTKP = cek_null(rs(13))
        Jumlah_Tanggungan = cek_null(rs(14))
        Nama_Jabatan = cleanStr(cek_null(rs(15)))
        WP_Luar_Negeri = cek_null(rs(16))
        Kode_Negara = cek_null(rs(17))
        Kode_Pajak = cek_null(rs(18))
        Jumlah_1 = cek_Money(rs(19))
        Jumlah_2 = cek_Money(rs(20))
        Jumlah_3 = cek_Money(rs(21))
        Jumlah_4 = cek_Money(rs(22))
        Jumlah_5 = cek_Money(rs(23))
        Jumlah_6 = cek_Money(rs(24))
        Jumlah_7 = cek_Money(rs(25))
        Jumlah_8 = cek_Money(rs(26))
        Jumlah_9 = cek_Money(rs(27))
        Jumlah_10 = cek_Money(rs(28))
        Jumlah_11 = cek_Money(rs(30))
        Jumlah_12 = cek_Money(rs(31))
        Jumlah_13 = cek_Money(rs(32))
        Jumlah_14 = cek_Money(rs(33))
        Jumlah_15 = cek_Money(rs(34))
        Jumlah_16 = cek_Money(rs(35))
        Jumlah_17 = cek_Money(rs(36))
        Jumlah_18 = cek_Money(rs(37))
        Jumlah_19 = cek_Money(rs(37))
        Jumlah_20 = cek_Money(rs(38))
        Status_Pindah = cek_null(rs(39))
        NPWP_Pemotong = cek_null(rs(40))
        Nama_Pemotong = cleanStr(cek_null(rs(41)))
        Tanggal_Bukti_Potong = cek_Date(rs(42))
        email = cleanStr(rs(43))
        
        If CStr(Year(Tanggal_Bukti_Potong)) = Tahun_Pajak Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Tahun Pajak tidak valid")
        End If
        
        If adddigit(CLng(Month(Tanggal_Bukti_Potong)), 2) = adddigit(cek_Lng(Masa_Pajak), 2) Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Masa Pajak tidak valid")
        End If
        
        kode_divisi = get_kode_combo(Me.cb_divisi, "-")
        '------------------
     
        If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
            data_Valid = False
            Call setListInfo(List1, "Data ke " & c & " NPWP_KPP tidak valid")
        ElseIf Trim(Nomor_Bukti_Potong) = "" Then
            data_Valid = False
            Call setListInfo(List1, "Data ke " & c & " Nomor bukti potong tidak valid")
        ElseIf Jumlah_1 + Jumlah_7 + Jumlah_14 + Jumlah_20 = 0 Then
            data_Valid = False
            Call setListInfo(List1, "Data ke " & c & " Nilai 0")
        ElseIf tbMkaryawan_isDataAda(NIK, npwp, nama) = False Then
            data_Valid = False
            Call setListInfo(List1, "Data ke " & c & " NIK / NPWP / Nama tidak ada di master karyawan")
        End If
    
        If data_Valid = True Then
        
            return1 = tbPph21Tahunan_insert(npwp_kpp, Masa_Pajak, Tahun_Pajak, Pembetulan, Nomor_Bukti_Potong, _
                                Masa_Perolehan_Awal, Masa_Perolehan_Akhir, npwp, NIK, nama, alamat, jenis_kelamin, _
                                Status_PTKP, Jumlah_Tanggungan, Nama_Jabatan, WP_Luar_Negeri, Kode_Negara, Kode_Pajak, _
                                Jumlah_1, Jumlah_2, Jumlah_3, Jumlah_4, Jumlah_5, Jumlah_6, Jumlah_7, Jumlah_8, _
                                Jumlah_9, Jumlah_10, Jumlah_11, Jumlah_12, Jumlah_13, Jumlah_14, Jumlah_15, _
                                Jumlah_16, Jumlah_17, Jumlah_18, Jumlah_19, Jumlah_20, Status_Pindah, NPWP_Pemotong, _
                                Nama_Pemotong, Tanggal_Bukti_Potong, kode_divisi, kd_proyek, email)
            If return1 = 1 Then
                jml_Insert = jml_Insert + 1
            ElseIf return1 = 2 Then
                jml_Update = jml_Update + 1
            ElseIf return1 = 3 Then
                jml_Skip = jml_Skip + 1
            Else
                Call pesan2("Insert error", , vbYellow)
                Exit Sub
            End If
        End If
        rs.MoveNext
    Loop
    MsgBox "Proses Import Selesai. Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update & " Jml Skip: " & jml_Skip, vbInformation

End Sub


Sub import_pph22()
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long
  
    Dim npwp_kpp As String, k02 As String, Masa_Pajak As String, Tahun_Pajak As String
    Dim Pembetulan As String, npwp As String, Nama_NPWP As String, alamat As String
    Dim Nomor_Bukti_Potong As String, Tanggal_Bukti_Potong As Date, k35 As String
    Dim k36 As String, k37 As String, k38 As String, k39 As String, k40 As String
    Dim k41 As String, k42 As String, k43 As String, Nilai_DPP As Currency, Tarif As String
    Dim Nilai_PPh As Currency, k47 As String, k48 As String, k49 As String, k50 As String
    Dim j51 As String, j52 As String, kode_divisi As String
    Dim kd_proyek As String, nott As String, nofaktur As String
                    
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
        npwp_kpp = cleanNpwp(rs(0))
        kd_proyek = cleanNpwp(rs(1))
        nott = cleanNpwp(rs(2))
        nofaktur = cleanNpwp(rs(3))
        
        k02 = cek_null(rs(4))
        
        If Trim(k02) = "F113304A" Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Kode Form tidak valid")
        End If
        
        Masa_Pajak = cek_null(rs(5))
        Tahun_Pajak = cek_null(rs(6))
        Pembetulan = cek_null(rs(7))
        npwp = cek_null(rs(8))
        
        If checkNPWP(npwp) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP tidak valid")
        End If
        
        Nama_NPWP = cleanStr(cek_null(rs(9)))
        alamat = cleanStr(cek_null(rs(10)))
        
        Call cek_npwpWP(npwp, Nama_NPWP, alamat)
        
        Nomor_Bukti_Potong = cek_null(rs(11))
        Tanggal_Bukti_Potong = cek_Date(rs(12))
        '-----------
        If CStr(Year(Tanggal_Bukti_Potong)) = Tahun_Pajak Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Tahun Pajak tidak valid")
        End If
        
        If adddigit(CLng(Month(Tanggal_Bukti_Potong)), 2) = adddigit(cek_Lng(Masa_Pajak), 2) Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Masa Pajak tidak valid")
        End If
        '-----------
        
        k35 = cek_null(rs(37))
        k36 = cek_null(rs(38))
        
        k37 = cek_null(rs(38))
        k38 = cek_null(rs(40))
        k39 = cek_null(rs(41))
        k40 = cek_null(rs(42))
        k41 = cek_null(rs(43))
        
        k42 = cek_null(rs(44))
        k43 = cek_null(rs(45))
        Nilai_DPP = cek_Money(rs(46))
        Tarif = cek_null(rs(47))
        Nilai_PPh = cek_Money(rs(48))
        
        k47 = cek_null(rs(49))
        k48 = cek_null(rs(50))
        k49 = cek_null(rs(51))
        k50 = cek_null(rs(52))
        j51 = cek_Money(rs(53))
        
        j52 = cek_null(rs(54))
        
        kode_divisi = get_kode_combo(Me.cb_divisi, "-")
        '------------------
     
        If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " NPWP_KPP tidak valid"
        ElseIf Trim(Nomor_Bukti_Potong) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nomor bukti potong tidak valid"
        ElseIf Nilai_PPh = 0 Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nilai 0"
        End If
    
        If data_Valid = True Then
        
            return1 = tbPph22_insert(npwp_kpp, k02, Masa_Pajak, Tahun_Pajak, Pembetulan, npwp, Nama_NPWP, _
                                    alamat, Nomor_Bukti_Potong, Tanggal_Bukti_Potong, k35, k36, k37, _
                                    k38, k39, k40, k41, k42, k43, Nilai_DPP, Tarif, Nilai_PPh, k47, _
                                    k48, k49, k50, j51, j52, kode_divisi, kd_proyek, nott, nofaktur)
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


Sub import_pph26()
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long
  
    Dim npwp_kpp As String, Kode_Form As String, Masa_Pajak As String, Tahun_Pajak As String
    Dim Pembetulan As String, npwp_wp As String, Nama_WP As String, Alamat_WP As String
    Dim Nomor_Bukti_Potong As String, Tanggal_Bukti_Potong As Date
    Dim Nilai_Bruto_1 As Currency, Tarif_1 As String, PPh_Yang_Dipotong__1 As Currency
    Dim Nilai_Bruto_2 As Currency, Tarif_2 As String, PPh_Yang_Dipotong__2 As Currency
    Dim Nilai_Bruto_3 As Currency, Tarif_3 As String, PPh_Yang_Dipotong__3 As Currency
    Dim Nilai_Bruto_4 As Currency, Tarif_4 As String, PPh_Yang_Dipotong__4 As Currency
    Dim Nilai_Bruto_5 As Currency, Tarif_5 As String, PPh_Yang_Dipotong__5 As Currency
    Dim Nilai_Bruto_6a As Currency, Tarif_6a As String, PPh_Yang_Dipotong__6a As Currency
    Dim Nilai_Bruto_6b As Currency, Tarif_6b As String, PPh_Yang_Dipotong__6b As Currency
    Dim Nilai_Bruto_6c As Currency, Tarif_6c As String, PPh_Yang_Dipotong__6c As Currency
    Dim Kode_Jasa_6d1 As String, Nilai_Bruto_6d1 As Currency, Tarif_6d1 As String, PPh_Yang_Dipotong__6d1 As Currency
    Dim Jumlah_Nilai_Bruto_ As Currency, Jumlah_PPh_Yang_Dipotong As Currency, kode_divisi As String
    Dim kd_proyek As String, nott As String, nofaktur As String, email As String
                    
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
        npwp_kpp = cleanNpwp(rs(0))
        kd_proyek = cleanNpwp(rs(1))
        nott = cleanNpwp(rs(2))
        nofaktur = cleanNpwp(rs(3))
                
        Kode_Form = cleanStr(rs(4))
        If Kode_Form = "F113308" Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Kode Form tidak valid")
        End If
        
        
        Masa_Pajak = cleanStr(rs(5))
        Tahun_Pajak = cleanStr(rs(6))
        Pembetulan = cleanStr(rs(7))
        npwp_wp = cleanStr(rs(8))
        
        If checkNPWP(npwp_wp) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP tidak valid")
        End If
        
        Nama_WP = cleanStr(rs(9))
        Alamat_WP = cleanStr(rs(10))
        
        Call cek_npwpWP(npwp_wp, Nama_WP, Alamat_WP)
        
        Nomor_Bukti_Potong = cleanStr(rs(11))
        Tanggal_Bukti_Potong = cek_Date(rs(12))
        
        '-----------
        If CStr(Year(Tanggal_Bukti_Potong)) = Tahun_Pajak Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Tahun Pajak tidak valid")
        End If
        
        If adddigit(CLng(Month(Tanggal_Bukti_Potong)), 2) = adddigit(cek_Lng(Masa_Pajak), 2) Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Masa Pajak tidak valid")
        End If
        '-----------
        
        Nilai_Bruto_1 = cek_Money(rs(13))
        Tarif_1 = cleanStr(rs(14))
        PPh_Yang_Dipotong__1 = cek_Money(rs(15))
        
        Nilai_Bruto_2 = cek_Money(rs(16))
        Tarif_2 = cleanStr(rs(17))
        PPh_Yang_Dipotong__2 = cek_Money(rs(18))
        
        Nilai_Bruto_3 = cek_Money(rs(19))
        Tarif_3 = cleanStr(rs(20))
        PPh_Yang_Dipotong__3 = cek_Money(rs(21))
        
        Nilai_Bruto_4 = cek_Money(rs(22))
        Tarif_4 = cleanStr(rs(23))
        PPh_Yang_Dipotong__4 = cek_Money(rs(24))
        
        Nilai_Bruto_5 = cek_Money(rs(25))
        Tarif_5 = cleanStr(rs(26))
        PPh_Yang_Dipotong__5 = cek_Money(rs(27))

        Nilai_Bruto_6a = cek_Money(rs(28))
        Tarif_6a = cleanStr(rs(29))
        PPh_Yang_Dipotong__6a = cek_Money(rs(30))
        
        Nilai_Bruto_6b = cek_Money(rs(31))
        Tarif_6b = cleanStr(rs(32))
        PPh_Yang_Dipotong__6b = cek_Money(rs(33))
        
        Nilai_Bruto_6c = cek_Money(rs(34))
        Tarif_6c = cleanStr(rs(35))
        PPh_Yang_Dipotong__6c = cek_Money(rs(36))
        
        
        Kode_Jasa_6d1 = cleanStr(rs(55))
        Nilai_Bruto_6d1 = cek_Money(rs(56))
        Tarif_6d1 = cleanStr(rs(57))
        PPh_Yang_Dipotong__6d1 = cek_Money(rs(58))
        
        Jumlah_Nilai_Bruto_ = cek_Money(rs(79))
        Jumlah_PPh_Yang_Dipotong = cek_Money(rs(80))
        email = cleanStr(rs(81))
        
        kode_divisi = get_kode_combo(Me.cb_divisi, "-")
        '------------------
     
        If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " NPWP_KPP tidak valid"
        ElseIf Trim(Nomor_Bukti_Potong) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nomor bukti potong tidak valid"
        ElseIf Jumlah_Nilai_Bruto_ + Jumlah_PPh_Yang_Dipotong = 0 Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nilai PPH 0"
        End If
    
        If data_Valid = True Then
        
            return1 = tbPph26_insert(npwp_kpp, Kode_Form, Masa_Pajak, Tahun_Pajak, Pembetulan, npwp_wp, Nama_WP, _
                                Alamat_WP, Nomor_Bukti_Potong, Tanggal_Bukti_Potong, _
                                Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1, _
                                Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, _
                                Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, _
                                Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, _
                                Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, _
                                Nilai_Bruto_6a, Tarif_6a, PPh_Yang_Dipotong__6a, _
                                Nilai_Bruto_6b, Tarif_6b, PPh_Yang_Dipotong__6b, _
                                Nilai_Bruto_6c, Tarif_6c, PPh_Yang_Dipotong__6c, _
                                Kode_Jasa_6d1, Nilai_Bruto_6d1, Tarif_6d1, PPh_Yang_Dipotong__6d1, _
                                Jumlah_Nilai_Bruto_, Jumlah_PPh_Yang_Dipotong, kode_divisi, _
                                kd_proyek, nott, nofaktur, email)
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


Sub import_pph42_Konstruksi()
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long
  
    Dim npwp_kpp As String, Kode_Form As String, Masa_Pajak As String
    Dim Tahun_Pajak As String, Pembetulan As String, npwp_wp As String
    Dim Nama_WP As String, Alamat_WP As String, Nomor_Bukti_Potong As String
    Dim Tanggal_Bukti_Potong As Date, Jenis_Hadiah_Undian_1 As String
    Dim Kode_Option_Tempat_Penyimpanan_1 As String, Jumlah_Nilai_Bruto_1 As Currency
    Dim Tarif_1 As String, PPh_Yang_Dipotong__1 As Currency
    Dim Jenis_Hadiah_Undian_2 As String, Kode_Option_Tempat_Penyimpanan_2 As String
    Dim Jumlah_Nilai_Bruto_2 As Currency, Tarif_2 As String
    Dim PPh_Yang_Dipotong__2 As Currency, Jenis_Hadiah_Undian_3 As String
    Dim Kode_Option_Tempat_Penyimpanan_3 As String, Jumlah_Nilai_Bruto_3 As Currency
    Dim Tarif_3 As String, PPh_Yang_Dipotong__3 As Currency
    Dim Jenis_Hadiah_Undian_4 As String, Kode_Option_Tempat_Penyimpanan_4 As String
    Dim Jumlah_Nilai_Bruto_4 As Currency, Tarif_4 As String
    Dim PPh_Yang_Dipotong__4 As Currency, Jenis_Hadiah_Undian_5 As String
    Dim Kode_Option_Tempat_Penyimpanan_5 As String, Jumlah_Nilai_Bruto_5 As Currency
    Dim Tarif_5 As String, PPh_Yang_Dipotong__5 As Currency
    Dim Jenis_Hadiah_Undian_6 As String, Jumlah_Nilai_Bruto_6 As Currency
    Dim Tarif_6 As String, PPh_Yang_Dipotong__6 As Currency
    Dim Jumlah_Nilai_Bruto_7 As Currency, Tarif_7 As String
    Dim PPh_Yang_Dipotong_7 As Currency, Jenis_Penghasilan_8 As String
    Dim Jumlah_Nilai_Bruto_8 As Currency, Tarif_8 As String
    Dim PPh_Yang_Dipotong_8 As Currency, Jumlah_PPh_Yang_Dipotong As Currency
    Dim Tanggal_Jatuh_Tempo_Obligasi As String, Tanggal_Perolehan_Obligasi As String
    Dim Tanggal_Penjualan_Obligasi As String, Holding_Periode_Obligasi As String
    Dim Time_Periode_Obligasi As String, kode_divisi As String
    Dim kd_proyek As String, nott As String, nofaktur As String, email As String
                    
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
        npwp_kpp = cleanNpwp(rs(0))
        kd_proyek = cleanNpwp(rs(1))
        nott = cleanNpwp(rs(2))
        nofaktur = cleanNpwp(rs(3))
        
        Kode_Form = cek_null(rs(4))
        
        If (Kode_Form = "F113316") Or (Kode_Form = "F113317") Or (Kode_Form = "F113319") Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Kode Form tidak valid")
        End If
        
        Masa_Pajak = cek_null(rs(5))
        Tahun_Pajak = cek_null(rs(6))
        Pembetulan = cek_null(rs(7))
        npwp_wp = cek_null(rs(8))
        
        If checkNPWP(npwp_wp) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP tidak valid")
        End If
        
        Nama_WP = cleanStr(cek_null(rs(9)))
        Alamat_WP = cleanStr(cek_null(rs(10)))
        
        Call cek_npwpWP(npwp_wp, Nama_WP, Alamat_WP)
        
        Nomor_Bukti_Potong = cek_null(rs(11))
        Tanggal_Bukti_Potong = cek_Date(rs(12))
        
        '-----------
        If CStr(Year(Tanggal_Bukti_Potong)) = Tahun_Pajak Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Tahun Pajak tidak valid")
        End If
        
        If adddigit(CLng(Month(Tanggal_Bukti_Potong)), 2) = adddigit(cek_Lng(Masa_Pajak), 2) Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Masa Pajak tidak valid")
        End If
        '-----------
        
        Jenis_Hadiah_Undian_1 = cek_null(rs(13))
        
        Kode_Option_Tempat_Penyimpanan_1 = cek_null(rs(14))
        Jumlah_Nilai_Bruto_1 = cek_Money(rs(15))
        Tarif_1 = cek_null(rs(16))
        PPh_Yang_Dipotong__1 = cek_Money(rs(17))
        Jenis_Hadiah_Undian_2 = cek_null(rs(18))
        
        Kode_Option_Tempat_Penyimpanan_2 = cek_null(rs(19))
        Jumlah_Nilai_Bruto_2 = cek_Money(rs(20))
        Tarif_2 = cek_null(rs(21))
        PPh_Yang_Dipotong__2 = cek_Money(rs(22))
        Jenis_Hadiah_Undian_3 = cek_null(rs(23))
        
        Kode_Option_Tempat_Penyimpanan_3 = cek_null(rs(24))
        Jumlah_Nilai_Bruto_3 = cek_Money(rs(25))
        Tarif_3 = cek_null(rs(26))
        PPh_Yang_Dipotong__3 = cek_Money(rs(27))
        Jenis_Hadiah_Undian_4 = cek_null(rs(28))
        
        Kode_Option_Tempat_Penyimpanan_4 = cek_null(rs(29))
        Jumlah_Nilai_Bruto_4 = cek_Money(rs(30))
        Tarif_4 = cek_null(rs(31))
        PPh_Yang_Dipotong__4 = cek_Money(rs(32))
        Jenis_Hadiah_Undian_5 = cek_null(rs(33))
        
        Kode_Option_Tempat_Penyimpanan_5 = cek_null(rs(34))
        Jumlah_Nilai_Bruto_5 = cek_Money(rs(35))
        Tarif_5 = cek_null(rs(36))
        PPh_Yang_Dipotong__5 = cek_Money(rs(37))
        Jenis_Hadiah_Undian_6 = cek_null(rs(38))
        
        Jumlah_Nilai_Bruto_6 = cek_Money(rs(39))
        Tarif_6 = cek_null(rs(40))
        PPh_Yang_Dipotong__6 = cek_Money(rs(41))
        Jumlah_Nilai_Bruto_7 = cek_Money(rs(42))
        Tarif_7 = cek_null(rs(43))
        
        PPh_Yang_Dipotong_7 = cek_Money(rs(44))
        Jenis_Penghasilan_8 = cek_null(rs(45))
        Jumlah_Nilai_Bruto_8 = cek_Money(rs(46))
        Tarif_8 = cek_null(rs(47))
        PPh_Yang_Dipotong_8 = cek_Money(rs(48))
        
        Jumlah_PPh_Yang_Dipotong = cek_Money(rs(49))
        Tanggal_Jatuh_Tempo_Obligasi = cek_null(rs(50))
        Tanggal_Perolehan_Obligasi = cek_null(rs(51))
        Tanggal_Penjualan_Obligasi = cek_null(rs(52))
        Holding_Periode_Obligasi = cek_null(rs(53))
        
        Time_Periode_Obligasi = cek_null(rs(54))
        email = cleanStr(rs(55))
        
        kode_divisi = get_kode_combo(Me.cb_divisi, "-")
        '------------------
     
        If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " NPWP_KPP tidak valid"
        ElseIf Trim(Nomor_Bukti_Potong) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nomor bukti potong tidak valid"
        ElseIf Jumlah_Nilai_Bruto_6 + Jumlah_PPh_Yang_Dipotong + PPh_Yang_Dipotong__1 + PPh_Yang_Dipotong__2 + PPh_Yang_Dipotong__3 + _
                PPh_Yang_Dipotong__4 + PPh_Yang_Dipotong__5 + PPh_Yang_Dipotong__6 = 0 Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nilai PPH 0"
        End If
    
        If data_Valid = True Then
        
            return1 = tbPph42Konstruksi_insert(npwp_kpp, Kode_Form, Masa_Pajak, Tahun_Pajak, Pembetulan, npwp_wp, _
                                                Nama_WP, Alamat_WP, Nomor_Bukti_Potong, Tanggal_Bukti_Potong, _
                                                Jenis_Hadiah_Undian_1, Kode_Option_Tempat_Penyimpanan_1, Jumlah_Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1, _
                                                Jenis_Hadiah_Undian_2, Kode_Option_Tempat_Penyimpanan_2, Jumlah_Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, _
                                                Jenis_Hadiah_Undian_3, Kode_Option_Tempat_Penyimpanan_3, Jumlah_Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, _
                                                Jenis_Hadiah_Undian_4, Kode_Option_Tempat_Penyimpanan_4, Jumlah_Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, _
                                                Jenis_Hadiah_Undian_5, Kode_Option_Tempat_Penyimpanan_5, Jumlah_Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, _
                                                Jenis_Hadiah_Undian_6, Jumlah_Nilai_Bruto_6, Tarif_6, PPh_Yang_Dipotong__6, _
                                                Jumlah_Nilai_Bruto_7, Tarif_7, PPh_Yang_Dipotong_7, Jenis_Penghasilan_8, _
                                                Jumlah_Nilai_Bruto_8, Tarif_8, PPh_Yang_Dipotong_8, Jumlah_PPh_Yang_Dipotong, _
                                                Tanggal_Jatuh_Tempo_Obligasi, Tanggal_Perolehan_Obligasi, _
                                                Tanggal_Penjualan_Obligasi, Holding_Periode_Obligasi, _
                                                Time_Periode_Obligasi, kode_divisi, kd_proyek, nott, nofaktur, email)
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


Sub import_pph42_Sewa()
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long
  
    Dim npwp_kpp As String, Kode_Form As String, Masa_Pajak As String
    Dim Tahun_Pajak As String, Pembetulan As String, npwp_wp As String
    Dim Nama_WP As String, Alamat_WP As String, Nomor_Bukti_Potong As String
    Dim Tanggal_Bukti_Potong As Date, Jenis_Hadiah_Undian_1 As String
    Dim Kode_Option_Tempat_Penyimpanan_1 As String, Jumlah_Nilai_Bruto_1 As Currency
    Dim Tarif_1 As String, PPh_Yang_Dipotong__1 As Currency
    Dim Jenis_Hadiah_Undian_2 As String, Kode_Option_Tempat_Penyimpanan_2 As String
    Dim Jumlah_Nilai_Bruto_2 As Currency, Tarif_2 As String
    Dim PPh_Yang_Dipotong__2 As Currency, Jenis_Hadiah_Undian_3 As String
    Dim Kode_Option_Tempat_Penyimpanan_3 As String, Jumlah_Nilai_Bruto_3 As Currency
    Dim Tarif_3 As String, PPh_Yang_Dipotong__3 As Currency
    Dim Jenis_Hadiah_Undian_4 As String, Kode_Option_Tempat_Penyimpanan_4 As String
    Dim Jumlah_Nilai_Bruto_4 As Currency, Tarif_4 As String
    Dim PPh_Yang_Dipotong__4 As Currency, Jenis_Hadiah_Undian_5 As String
    Dim Kode_Option_Tempat_Penyimpanan_5 As String, Jumlah_Nilai_Bruto_5 As Currency
    Dim Tarif_5 As String, PPh_Yang_Dipotong__5 As Currency
    Dim Jenis_Hadiah_Undian_6 As String, Jumlah_Nilai_Bruto_6 As Currency
    Dim Tarif_6 As String, PPh_Yang_Dipotong__6 As Currency
    Dim Jumlah_Nilai_Bruto_7 As Currency, Tarif_7 As String
    Dim PPh_Yang_Dipotong_7 As Currency, Jenis_Penghasilan_8 As String
    Dim Jumlah_Nilai_Bruto_8 As Currency, Tarif_8 As String
    Dim PPh_Yang_Dipotong_8 As Currency, Jumlah_PPh_Yang_Dipotong As Currency
    Dim Tanggal_Jatuh_Tempo_Obligasi As String, Tanggal_Perolehan_Obligasi As String
    Dim Tanggal_Penjualan_Obligasi As String, Holding_Periode_Obligasi As String
    Dim Time_Periode_Obligasi As String, kode_divisi As String
    Dim kd_proyek As String, nott As String, nofaktur As String, email As String
                    
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
        npwp_kpp = cleanNpwp(rs(0))
        kd_proyek = cleanNpwp(rs(1))
        nott = cleanNpwp(rs(2))
        nofaktur = cleanNpwp(rs(3))
        
        Kode_Form = cek_null(rs(4))
        If (Kode_Form = "F113312") Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Kode Form tidak valid")
        End If
        
        Masa_Pajak = cek_null(rs(5))
        Tahun_Pajak = cek_null(rs(6))
        Pembetulan = cek_null(rs(7))
        npwp_wp = cek_null(rs(8))
        
        If checkNPWP(npwp_wp) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP tidak valid")
        End If
        
        Nama_WP = cleanStr(cek_null(rs(9)))
        Alamat_WP = cleanStr(cek_null(rs(10)))
        
        Call cek_npwpWP(npwp_wp, Nama_WP, Alamat_WP)
        
        Nomor_Bukti_Potong = cek_null(rs(11))
        Tanggal_Bukti_Potong = cek_Date(rs(12))
        
        '-----------
        If CStr(Year(Tanggal_Bukti_Potong)) = Tahun_Pajak Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Tahun Pajak tidak valid")
        End If
        
        If adddigit(CLng(Month(Tanggal_Bukti_Potong)), 2) = adddigit(cek_Lng(Masa_Pajak), 2) Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Masa Pajak tidak valid")
        End If
        '-----------
        
        Jenis_Hadiah_Undian_1 = cek_null(rs(13))
        
        Kode_Option_Tempat_Penyimpanan_1 = cek_null(rs(14))
        Jumlah_Nilai_Bruto_1 = cek_Money(rs(15))
        Tarif_1 = cek_null(rs(16))
        PPh_Yang_Dipotong__1 = cek_Money(rs(17))
        Jenis_Hadiah_Undian_2 = cek_null(rs(18))
        
        Kode_Option_Tempat_Penyimpanan_2 = cek_null(rs(19))
        Jumlah_Nilai_Bruto_2 = cek_Money(rs(20))
        Tarif_2 = cek_null(rs(21))
        PPh_Yang_Dipotong__2 = cek_Money(rs(22))
        Jenis_Hadiah_Undian_3 = cek_null(rs(23))
        
        Kode_Option_Tempat_Penyimpanan_3 = cek_null(rs(24))
        Jumlah_Nilai_Bruto_3 = cek_Money(rs(25))
        Tarif_3 = cek_null(rs(26))
        PPh_Yang_Dipotong__3 = cek_Money(rs(27))
        Jenis_Hadiah_Undian_4 = cek_null(rs(28))
        
        Kode_Option_Tempat_Penyimpanan_4 = cek_null(rs(29))
        Jumlah_Nilai_Bruto_4 = cek_Money(rs(30))
        Tarif_4 = cek_null(rs(31))
        PPh_Yang_Dipotong__4 = cek_Money(rs(32))
        Jenis_Hadiah_Undian_5 = cek_null(rs(33))
        
        Kode_Option_Tempat_Penyimpanan_5 = cek_null(rs(34))
        Jumlah_Nilai_Bruto_5 = cek_Money(rs(35))
        Tarif_5 = cek_null(rs(36))
        PPh_Yang_Dipotong__5 = cek_Money(rs(37))
        Jenis_Hadiah_Undian_6 = cek_null(rs(38))
        
        Jumlah_Nilai_Bruto_6 = cek_Money(rs(39))
        Tarif_6 = cek_null(rs(40))
        PPh_Yang_Dipotong__6 = cek_Money(rs(41))
        Jumlah_Nilai_Bruto_7 = cek_Money(rs(42))
        Tarif_7 = cek_null(rs(43))
        
        PPh_Yang_Dipotong_7 = cek_Money(rs(44))
        Jenis_Penghasilan_8 = cek_null(rs(45))
        Jumlah_Nilai_Bruto_8 = cek_Money(rs(46))
        Tarif_8 = cek_null(rs(47))
        PPh_Yang_Dipotong_8 = cek_Money(rs(48))
        
        Jumlah_PPh_Yang_Dipotong = cek_Money(rs(49))
        Tanggal_Jatuh_Tempo_Obligasi = cek_null(rs(50))
        Tanggal_Perolehan_Obligasi = cek_null(rs(51))
        Tanggal_Penjualan_Obligasi = cek_null(rs(52))
        Holding_Periode_Obligasi = cek_null(rs(53))
        
        Time_Periode_Obligasi = cek_null(rs(54))
        email = cleanStr(rs(55))
        
        kode_divisi = get_kode_combo(Me.cb_divisi, "-")
        '------------------
     
        If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " NPWP_KPP tidak valid"
        ElseIf Trim(Nomor_Bukti_Potong) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nomor bukti potong tidak valid"
        ElseIf Jumlah_PPh_Yang_Dipotong + PPh_Yang_Dipotong__1 + PPh_Yang_Dipotong__2 + PPh_Yang_Dipotong__3 + _
                PPh_Yang_Dipotong__4 + PPh_Yang_Dipotong__5 + PPh_Yang_Dipotong__6 + Jumlah_PPh_Yang_Dipotong + Jumlah_Nilai_Bruto_5 + Jumlah_Nilai_Bruto_6 = 0 Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nilai PPH 0"
        End If
    
        If data_Valid = True Then
        
            return1 = tbPph42Sewa_insert(npwp_kpp, Kode_Form, Masa_Pajak, Tahun_Pajak, Pembetulan, npwp_wp, _
                                                Nama_WP, Alamat_WP, Nomor_Bukti_Potong, Tanggal_Bukti_Potong, _
                                                Jenis_Hadiah_Undian_1, Kode_Option_Tempat_Penyimpanan_1, Jumlah_Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1, _
                                                Jenis_Hadiah_Undian_2, Kode_Option_Tempat_Penyimpanan_2, Jumlah_Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, _
                                                Jenis_Hadiah_Undian_3, Kode_Option_Tempat_Penyimpanan_3, Jumlah_Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, _
                                                Jenis_Hadiah_Undian_4, Kode_Option_Tempat_Penyimpanan_4, Jumlah_Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, _
                                                Jenis_Hadiah_Undian_5, Kode_Option_Tempat_Penyimpanan_5, Jumlah_Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, _
                                                Jenis_Hadiah_Undian_6, Jumlah_Nilai_Bruto_6, Tarif_6, PPh_Yang_Dipotong__6, _
                                                Jumlah_Nilai_Bruto_7, Tarif_7, PPh_Yang_Dipotong_7, Jenis_Penghasilan_8, _
                                                Jumlah_Nilai_Bruto_8, Tarif_8, PPh_Yang_Dipotong_8, Jumlah_PPh_Yang_Dipotong, _
                                                Tanggal_Jatuh_Tempo_Obligasi, Tanggal_Perolehan_Obligasi, _
                                                Tanggal_Penjualan_Obligasi, Holding_Periode_Obligasi, _
                                                Time_Periode_Obligasi, kode_divisi, kd_proyek, nott, nofaktur, email)
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

Sub import_pph42_obligasi()
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long
  
    Dim npwp_kpp As String, Kode_Form As String, Masa_Pajak As String
    Dim Tahun_Pajak As String, Pembetulan As String, npwp_wp As String
    Dim Nama_WP As String, Alamat_WP As String, Nomor_Bukti_Potong As String
    Dim Tanggal_Bukti_Potong As Date, Jenis_Hadiah_Undian_1 As String
    Dim Kode_Option_Tempat_Penyimpanan_1 As String, Jumlah_Nilai_Bruto_1 As Currency
    Dim Tarif_1 As String, PPh_Yang_Dipotong__1 As Currency
    Dim Jenis_Hadiah_Undian_2 As String, Kode_Option_Tempat_Penyimpanan_2 As String
    Dim Jumlah_Nilai_Bruto_2 As Currency, Tarif_2 As String
    Dim PPh_Yang_Dipotong__2 As Currency, Jenis_Hadiah_Undian_3 As String
    Dim Kode_Option_Tempat_Penyimpanan_3 As String, Jumlah_Nilai_Bruto_3 As Currency
    Dim Tarif_3 As String, PPh_Yang_Dipotong__3 As Currency
    Dim Jenis_Hadiah_Undian_4 As String, Kode_Option_Tempat_Penyimpanan_4 As String
    Dim Jumlah_Nilai_Bruto_4 As Currency, Tarif_4 As String
    Dim PPh_Yang_Dipotong__4 As Currency, Jenis_Hadiah_Undian_5 As String
    Dim Kode_Option_Tempat_Penyimpanan_5 As String, Jumlah_Nilai_Bruto_5 As Currency
    Dim Tarif_5 As String, PPh_Yang_Dipotong__5 As Currency
    Dim Jenis_Hadiah_Undian_6 As String, Jumlah_Nilai_Bruto_6 As Currency
    Dim Tarif_6 As String, PPh_Yang_Dipotong__6 As Currency
    Dim Jumlah_Nilai_Bruto_7 As Currency, Tarif_7 As String
    Dim PPh_Yang_Dipotong_7 As Currency, Jenis_Penghasilan_8 As String
    Dim Jumlah_Nilai_Bruto_8 As Currency, Tarif_8 As String
    Dim PPh_Yang_Dipotong_8 As Currency, Jumlah_PPh_Yang_Dipotong As Currency
    Dim Tanggal_Jatuh_Tempo_Obligasi As String, Tanggal_Perolehan_Obligasi As String
    Dim Tanggal_Penjualan_Obligasi As String, Holding_Periode_Obligasi As String
    Dim Time_Periode_Obligasi As String, kode_divisi As String
    Dim kd_proyek As String, nott As String, nofaktur As String, email As String
                    
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
        npwp_kpp = cleanNpwp(rs(0))
        kd_proyek = cleanNpwp(rs(1))
        nott = cleanNpwp(rs(2))
        nofaktur = cleanNpwp(rs(3))
        
        Kode_Form = cek_null(rs(4))
        If (Kode_Form = "F113317") Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Kode Form tidak valid")
        End If
        
        Masa_Pajak = cek_null(rs(5))
        Tahun_Pajak = cek_null(rs(6))
        Pembetulan = cek_null(rs(7))
        npwp_wp = cek_null(rs(8))
        
        If checkNPWP(npwp_wp) = False Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NPWP tidak valid")
        End If
        
        Nama_WP = cleanStr(cek_null(rs(9)))
        Alamat_WP = cleanStr(cek_null(rs(10)))
        
        Call cek_npwpWP(npwp_wp, Nama_WP, Alamat_WP)
        
        Nomor_Bukti_Potong = cek_null(rs(11))
        Tanggal_Bukti_Potong = cek_Date(rs(12))
        
        '-----------
        If CStr(Year(Tanggal_Bukti_Potong)) = Tahun_Pajak Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Tahun Pajak tidak valid")
        End If
        
        If adddigit(CLng(Month(Tanggal_Bukti_Potong)), 2) = adddigit(cek_Lng(Masa_Pajak), 2) Then
        Else
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Masa Pajak tidak valid")
        End If
        '-----------
        
        Jenis_Hadiah_Undian_1 = cek_null(rs(13))
        
        Kode_Option_Tempat_Penyimpanan_1 = cek_null(rs(14))
        Jumlah_Nilai_Bruto_1 = cek_Money(rs(15))
        Tarif_1 = cek_null(rs(16))
        PPh_Yang_Dipotong__1 = cek_Money(rs(17))
        Jenis_Hadiah_Undian_2 = cek_null(rs(18))
        
        Kode_Option_Tempat_Penyimpanan_2 = cek_null(rs(19))
        Jumlah_Nilai_Bruto_2 = cek_Money(rs(20))
        Tarif_2 = cek_null(rs(21))
        PPh_Yang_Dipotong__2 = cek_Money(rs(22))
        Jenis_Hadiah_Undian_3 = cek_null(rs(23))
        
        Kode_Option_Tempat_Penyimpanan_3 = cek_null(rs(24))
        Jumlah_Nilai_Bruto_3 = cek_Money(rs(25))
        Tarif_3 = cek_null(rs(26))
        PPh_Yang_Dipotong__3 = cek_Money(rs(27))
        Jenis_Hadiah_Undian_4 = cek_null(rs(28))
        
        Kode_Option_Tempat_Penyimpanan_4 = cek_null(rs(29))
        Jumlah_Nilai_Bruto_4 = cek_Money(rs(30))
        Tarif_4 = cek_null(rs(31))
        PPh_Yang_Dipotong__4 = cek_Money(rs(32))
        Jenis_Hadiah_Undian_5 = cek_null(rs(33))
        
        Kode_Option_Tempat_Penyimpanan_5 = cek_null(rs(34))
        Jumlah_Nilai_Bruto_5 = cek_Money(rs(35))
        Tarif_5 = cek_null(rs(36))
        PPh_Yang_Dipotong__5 = cek_Money(rs(37))
        Jenis_Hadiah_Undian_6 = cek_null(rs(38))
        
        Jumlah_Nilai_Bruto_6 = cek_Money(rs(39))
        Tarif_6 = cek_null(rs(40))
        PPh_Yang_Dipotong__6 = cek_Money(rs(41))
        Jumlah_Nilai_Bruto_7 = cek_Money(rs(42))
        Tarif_7 = cek_null(rs(43))
        
        PPh_Yang_Dipotong_7 = cek_Money(rs(44))
        Jenis_Penghasilan_8 = cek_null(rs(45))
        Jumlah_Nilai_Bruto_8 = cek_Money(rs(46))
        Tarif_8 = cek_null(rs(47))
        PPh_Yang_Dipotong_8 = cek_Money(rs(48))
        
        Jumlah_PPh_Yang_Dipotong = cek_Money(rs(49))
        Tanggal_Jatuh_Tempo_Obligasi = cek_null(rs(50))
        Tanggal_Perolehan_Obligasi = cek_null(rs(51))
        Tanggal_Penjualan_Obligasi = cek_null(rs(52))
        Holding_Periode_Obligasi = cek_null(rs(53))
        
        Time_Periode_Obligasi = cek_null(rs(54))
        email = cleanStr(rs(55))
        
        kode_divisi = get_kode_combo(Me.cb_divisi, "-")
        '------------------
     
        If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " NPWP_KPP tidak valid"
        ElseIf Trim(Nomor_Bukti_Potong) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nomor bukti potong tidak valid"
        ElseIf Jumlah_PPh_Yang_Dipotong + PPh_Yang_Dipotong__1 + PPh_Yang_Dipotong__2 + PPh_Yang_Dipotong__3 + _
                PPh_Yang_Dipotong__4 + PPh_Yang_Dipotong__5 + PPh_Yang_Dipotong__6 + Jumlah_PPh_Yang_Dipotong + Jumlah_Nilai_Bruto_5 + Jumlah_Nilai_Bruto_6 = 0 Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nilai PPH 0"
        End If
    
        If data_Valid = True Then
        
            return1 = tbPph42Obligasi_insert(npwp_kpp, Kode_Form, Masa_Pajak, Tahun_Pajak, Pembetulan, npwp_wp, _
                                                Nama_WP, Alamat_WP, Nomor_Bukti_Potong, Tanggal_Bukti_Potong, _
                                                Jenis_Hadiah_Undian_1, Kode_Option_Tempat_Penyimpanan_1, Jumlah_Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1, _
                                                Jenis_Hadiah_Undian_2, Kode_Option_Tempat_Penyimpanan_2, Jumlah_Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, _
                                                Jenis_Hadiah_Undian_3, Kode_Option_Tempat_Penyimpanan_3, Jumlah_Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, _
                                                Jenis_Hadiah_Undian_4, Kode_Option_Tempat_Penyimpanan_4, Jumlah_Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, _
                                                Jenis_Hadiah_Undian_5, Kode_Option_Tempat_Penyimpanan_5, Jumlah_Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, _
                                                Jenis_Hadiah_Undian_6, Jumlah_Nilai_Bruto_6, Tarif_6, PPh_Yang_Dipotong__6, _
                                                Jumlah_Nilai_Bruto_7, Tarif_7, PPh_Yang_Dipotong_7, Jenis_Penghasilan_8, _
                                                Jumlah_Nilai_Bruto_8, Tarif_8, PPh_Yang_Dipotong_8, Jumlah_PPh_Yang_Dipotong, _
                                                Tanggal_Jatuh_Tempo_Obligasi, Tanggal_Perolehan_Obligasi, _
                                                Tanggal_Penjualan_Obligasi, Holding_Periode_Obligasi, _
                                                Time_Periode_Obligasi, kode_divisi, kd_proyek, nott, nofaktur, email)
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
