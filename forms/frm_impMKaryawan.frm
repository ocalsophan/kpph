VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_impMKaryawan 
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
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
      Left            =   117
      TabIndex        =   5
      ToolTipText     =   "Log Hasil Import File Excel. Double Klik untuk Simpan"
      Top             =   5616
      Width           =   12038
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   " Isi File "
      Height          =   3406
      Left            =   117
      TabIndex        =   7
      Top             =   2160
      Width           =   12038
      Begin VB.CommandButton cmd_import 
         Caption         =   "Import"
         Height          =   375
         Left            =   10560
         TabIndex        =   4
         Top             =   2808
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DGrid1 
         Height          =   2457
         Left            =   117
         TabIndex        =   8
         Top             =   234
         Width           =   11791
         _ExtentX        =   20796
         _ExtentY        =   4313
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
      Caption         =   " Pilih File Import "
      Height          =   1339
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   12038
      Begin VB.CommandButton cmd_info 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Info format file"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmd_browse 
         Caption         =   "Browse"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   10478
      End
   End
   Begin VB.Label lb_caption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Import Master Karyawan"
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
Attribute VB_Name = "frm_impMKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset

Private Sub cmd_browse_Click()
    Dim f As String
    Dim fl As Object
  
    On Error GoTo er1
    MsgBox "Salah Pilih Format akan menampilkan hasil yang salah", vbExclamation
    Me.disable_Form
    CD.InitDir = App.Path & "\Import\"
    CD.Filter = "Excel file (*.xls;*.xlsx)|*.xls;*.xlsx"
    CD.FileName = ""
    CD.ShowOpen
    f = CD.FileName
    
    If Trim(f) <> "" Then
        Me.Text1 = f
        If is_file_ada(f) = True Then
            'File Valid
            If open_xls_lateBinding(fl, f) <> 0 Then
                Call pesan2("error open EXCEL", , vbYellow)
            Else
                Call Load_Excel_2Rs(fl, 1, rs, Me.StatusBar1, 1, 1)
                Set Me.DGrid1.DataSource = rs
            End If
            Call close_xls_lateBinding(fl)
        Else
            MsgBox "File tidak valid", vbCritical
        End If
    End If
    MsgBox "Jumlah data di file : " & RecordCount(rs)
    Set Me.DGrid1.DataSource = rs
    Me.Enable_Form
    Exit Sub
er1:
    MsgBox Err.Description, vbCritical
    Me.Enable_Form
End Sub

Private Sub cmd_contoh_File_Click()
  Dim t As String
  
  Me.disable_Form
    t = "Header dimulai dari baris 3, data dimulai dari baris 4, kolom di mulai dari A" & vbCr & _
        "Susunan Kolom: " & vbCr & _
        "no1, KD_PROYEK, NM_PROYEK, TGL_INPUT, RAPK_NCL, RAPK_REGULER" & vbCr & _
        "Acuan adalah KD_PROYEK. Jika KD_PROYEK sudah ada, akan di update."
        MsgBox t, vbInformation
  Me.Enable_Form
End Sub

Sub disable_Form()
  Me.Frame2.Enabled = False
  Me.Frame3.Enabled = False
  Me.List1.Enabled = False
End Sub

Sub Enable_Form()
  Me.Frame2.Enabled = True
  Me.Frame3.Enabled = True
  Me.List1.Enabled = True
End Sub

Private Sub cmd_import_Click()
    'ambil semua data
  
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long, hasil As Integer, mod1 As Long
  
    Dim No1 As Long, nik As String, npwp As String
    Dim nama As String, alamat As String, jenis_kelamin As String
    Dim ptkp As String
  
    '--------------------------
    Dim t As String, ps, sql As String
    Dim data_Valid As Boolean
  
    On Error GoTo er1
    
    '---------------------
    If dbMySQL_open = True Then
    Else
        MsgBox "Open Database Tidak Sukses", vbExclamation
    End If
    '---------------------
  
    'konfirmasi,
    ps = MsgBox("Yakin akan import Data ?" & vbCr & "Pastikan Regional Setting: Indonesia", vbYesNo)
    If ps = vbNo Then Exit Sub
  
    jRec = RecordCount(rs)
    If jRec <= 0 Then
        MsgBox "Tidak ada data"
        Exit Sub
    End If
    
    Me.disable_Form
  
    rs.MoveFirst
    c = 0
    jml_Insert = 0
    jml_Update = 0
    Me.List1.Clear
    Do While rs.EOF = False
        c = c + 1
        Call info_progress(Me.StatusBar1, 1, c, jRec, "Run Import. Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update)
        data_Valid = True
        mod1 = c Mod 5000
        If mod1 = 0 Then Call dbMySQL_open
        
        'ssssssssssssssssss
        No1 = cleanStr(rs(0))
        nik = cleanStr(rs(1))
        npwp = cleanStr(rs(2))
        nama = cleanStr(rs(3))
        alamat = cleanStr(rs(4))
        jenis_kelamin = cleanStr(rs(5))
        
        If jenis_kelamin = "P" Or jenis_kelamin = "L" Then
        Else
            jenis_kelamin = "P"
        End If
        
        ptkp = cleanStr(rs(6))
     
        If Trim(nama) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nama tidak valid"
            Me.List1.ListIndex = Me.List1.ListCount - 1
        ElseIf Not (Trim(nama) <> "" And (Trim(npwp) <> "" Or Trim(nik) <> "")) Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Nama kosong atau NPWP/NIK Kosong"
            Me.List1.ListIndex = Me.List1.ListCount - 1
        ElseIf tbM_Ptkp_isDataAda(ptkp) = False Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " identifikasi PTKP tidak valid"
            Me.List1.ListIndex = Me.List1.ListCount - 1
        End If
    
        If data_Valid = True Then
            If tbMkaryawan_isDataAda(nik, npwp, nama) = True Then
                'update
                Me.List1.AddItem "Data ke " & c & " sudah ada. Skip"
            Else
                'insert
                If tbMkaryawan_insert(nik, npwp, nama, alamat, jenis_kelamin, ptkp) = True Then
                    jml_Insert = jml_Insert + 1
                Else
                    Call pesan2("Insert data ERROR", , vbYellow)
                    Exit Sub
                End If
            End If
        End If
        rs.MoveNext
    Loop
  
    MsgBox "Proses Import Selesai. Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update, vbInformation
    Me.Enable_Form
    Exit Sub
er1:
    MsgBox Err.Description, vbCritical
    Me.Enable_Form
End Sub


Function is_data_sudah_ada(kodeBank As String) As Boolean
  
  Dim sql As String, t As String
  
  sql = "select count(*) from tbank where kodebank = '" & CekPetik(kodeBank) & "'"
  t = cari_data1(cnn, sql, True)
  If CInt(t) > 0 Then
    is_data_sudah_ada = True
  Else
    is_data_sudah_ada = False
  End If
  
End Function



Private Sub cmd_info_Click()
    Dim t As String
    
    t = "data diletakkan di EXCEL, " & vbCr & _
        "dimulai dari A1 (header) " & vbCr & _
        "dengan susunan kolom:" & vbCr & _
        "No., NIK, NPWP, Nama, Alamat, " & vbCr & _
        "jenis Kelamin, PTKP " & vbCr & _
        "-- Data di mulai dari baris ke 2--"
    MsgBox t, vbInformation
End Sub

Private Sub Form_Load()
  Dim sql As String
  
  Me.Text1 = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call dbMySQL_open
    Call info(2, "cek biaya_jabatan", frMenu1.StatusBar1)
    Call tbPph21Tahunan2_setBiayaJabatanPerBulan
    Call info(2, "cek biaya_jabatan : OK", frMenu1.StatusBar1)
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
    namaFile = InputBox("Lokasi Simpan", , namaFile)
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
