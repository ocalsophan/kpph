VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Ebupotpph26_imp 
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
      Begin VB.TextBox txt_divisi 
         Height          =   375
         Left            =   10680
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   840
         Width           =   1095
      End
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kode Divisi"
         Height          =   210
         Left            =   9720
         TabIndex        =   10
         Top             =   915
         Width           =   795
      End
   End
   Begin VB.Label lb_caption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Import template ebupot26"
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
Attribute VB_Name = "frm_Ebupotpph26_imp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset

Private Sub cmd_browse_Click()
    Dim f As String
    Dim fl As Object
    Dim jmlKolom As Integer
  
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
    
    jmlKolom = rs.Fields.Count
    If jmlKolom <> 96 Then
        MsgBox "Jumlah kolom tidak valid", vbCritical
        Me.cmd_import.Enabled = False
    Else
        Call pesan2("jumlah kolom: " & jmlKolom, 1)
        Me.cmd_import.Enabled = True
    End If
    
    MsgBox "Jumlah data di file : " & RecordCount(rs)
    Set Me.DGrid1.DataSource = rs
    Me.Enable_Form
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
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
            
    Dim hapus_data As Boolean
    
    Dim Tahun_Pajak As String, No_Bukti_Akuntansi As String, NPWP_WP_yang_Dipotong As String
    Dim TIN_ As String, Masa_Pajak As String, Kode_Objek_Pajak As String
    Dim No_Faktur_Pajak As String, Tgl_Dokumen_ As String, Tanggal_Bukti_Potong As String
    Dim Kode_Proyek As String
    

    '--------------------------
    Dim t As String, ps, sql As String, ret1 As String
    Dim data_Valid As Boolean
    Dim rsParam As ADODB.Recordset, a As Integer
  
    On Error GoTo er1
    
    If Trim(Me.txt_divisi) = "" Then
        Call pesan2("kode divisi tidak valid. exit")
        Exit Sub
    ElseIf Len(Trim(Me.txt_divisi)) <> 6 Then
        Call pesan2("kode divisi tidak valid. exit")
        Exit Sub
    End If
    
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
    
    If createRS_duplicate(rs, rsParam) = False Then
        Call pesan2("Create rsParam fail.exit")
        Me.Enable_Form
        Exit Sub
    End If
    
    c = 0
    jml_Insert = 0
    jml_Update = 0
    Me.List1.Clear
    hapus_data = True
    Do While rs.EOF = False
        c = c + 1
        Call info(1, "Run Import " & c & " of " & jRec & ". Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update, Me.StatusBar1)
        data_Valid = True
        mod1 = c Mod 5000
        If mod1 = 0 Then Call dbMySQL_open
        
        
        Tahun_Pajak = cek_null(rs(8))
        Kode_Proyek = cek_null(rs(1))
        No_Bukti_Akuntansi = cek_null(rs(2))
        NPWP_WP_yang_Dipotong = cek_null(rs(10))
        TIN_ = cek_null(rs(11))
        Masa_Pajak = cek_null(rs(7))
        Kode_Objek_Pajak = cek_null(rs(17))
        No_Faktur_Pajak = cek_null(rs(5))
        Tgl_Dokumen_ = cek_null(rs(4))
        Tanggal_Bukti_Potong = cek_null(rs(20))
        
        'ssssssssssssssssss
        
        'If hapus_data = True Then
        '    Call pesan2("data tahun " & tahun & " akan di hapus terlebih dahulu")
        '    sql = "delete from all2016_fp where tahun = '" & tahun & "'"
        '    If ExecSQL1(cnn, sql) <> 0 Then
        '        sql = InputBox("sql error", "", sql)
        '    End If
        '    hapus_data = False
        'End If
     
     
     
        If Trim(Tahun_Pajak) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Tahun tidak valid"
            Me.List1.ListIndex = Me.List1.ListCount - 1
        ElseIf Trim(Kode_Proyek) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " kode proyek tidak valid"
            Me.List1.ListIndex = Me.List1.ListCount - 1
        'ElseIf Trim(No_Bukti_Akuntansi) = "" Then
        '    data_Valid = False
        '    Me.List1.AddItem "Data ke " & c & " No_Bukti_Akuntansi tidak valid"
        '    Me.List1.ListIndex = Me.List1.ListCount - 1
        'ElseIf Trim(NPWP_WP_yang_Dipotong) = "" Then
        '    data_Valid = False
        '    Me.List1.AddItem "Data ke " & c & " NPWP_WP_yang_Dipotong tidak valid"
        '    Me.List1.ListIndex = Me.List1.ListCount - 1
        'ElseIf Trim(NPWP_WP_yang_Dipotong) <> "" Then
        '    If Trim(NPWP_WP_yang_Dipotong) = "0" Then
        '    ElseIf Trim(NPWP_WP_yang_Dipotong) = "000000000000000" Or Trim(NPWP_WP_yang_Dipotong) = "0000000000000000" Then
        '    Else
        '        If checkNPWP(NPWP_WP_yang_Dipotong) = False Then
        '            data_Valid = False
        '            Me.List1.AddItem "Data ke " & c & " NPWP_WP_yang_Dipotong tidak valid"
        '            Me.List1.ListIndex = Me.List1.ListCount - 1
        '        End If
        '    End If
        ElseIf Trim(TIN_) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " TIN_ tidak valid"
            Me.List1.ListIndex = Me.List1.ListCount - 1
        ElseIf Trim(Masa_Pajak) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Masa_Pajak tidak valid"
            Me.List1.ListIndex = Me.List1.ListCount - 1
        ElseIf isKodeBuktiPotongValid(Kode_Objek_Pajak) = False Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Kode_Objek_Pajak tidak valid"
            Me.List1.ListIndex = Me.List1.ListCount - 1
        ElseIf Trim(No_Faktur_Pajak) = "" Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " No_Faktur_Pajak tidak valid"
            Me.List1.ListIndex = Me.List1.ListCount - 1
        ElseIf isddMMyyyy(Tgl_Dokumen_) = False Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Tgl_Dokumen_ tidak valid"
            Me.List1.ListIndex = Me.List1.ListCount - 1
        ElseIf isddMMyyyy(Tanggal_Bukti_Potong) = False Then
            data_Valid = False
            Me.List1.AddItem "Data ke " & c & " Tanggal_Bukti_Potong tidak valid"
            Me.List1.ListIndex = Me.List1.ListCount - 1
        End If
    
        If data_Valid = True Then
            'prepare data
            If rsParam.RecordCount > 0 Then
                rsParam.MoveFirst
                rsParam.Delete
            End If
            rsParam.AddNew
            For a = 0 To rs.Fields.Count - 1
                If rs(a).Type = adCurrency Then
                    rsParam.Fields(a).Value = cek_Money(rs.Fields(a).Value)
                Else
                    rsParam.Fields(a).Value = rs.Fields(a).Value
                End If
            Next
            rsParam.Update
        
        
            ret1 = tbebupot26_insert(rsParam, Trim(Me.txt_divisi), True)
            If ret1 = "" Then
                jml_Insert = jml_Insert + 1
            ElseIf ret1 = "update" Then
                Me.List1.AddItem "Data ke " & c & " sudah ada. Update"
                jml_Update = jml_Update + 1
                Me.List1.ListIndex = Me.List1.ListCount - 1
            Else
                Me.List1.AddItem "Data ke " & c & " error insert / update. err:" & _
                                    ret1
            End If
        End If
        rs.MoveNext
    Loop
  
    MsgBox "Proses Import Selesai. Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update, vbInformation
    Me.Enable_Form
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
    Me.Enable_Form
End Sub


Function isKodeBuktiPotongValid(kode As String) As Boolean
    Dim acuan, acuan2 As String
    
    acuan2 = "24-100-01,24-100-02,24-101-01,24-102-01,24-103-01,24-104-01, " & _
            "24-104-02, 24-104-03,24-104-04,24-104-05,24-104-06,24-104-07, " & _
            "24-104-08,24-104-09, 24-104-10,24-104-11,24-104-12,24-104-13, " & _
            "24-104-14,24-104-15,24-104-16, 24-104-17,24-104-18,24-104-19, " & _
            "24-104-20,24-104-21,24-104-22,24-104-23, 24-104-24,24-104-25, " & _
            "24-104-26,24-104-27,24-104-28,24-104-29,24-104-30, 24-104-31, " & _
            "24-104-32,24-104-33,24-104-34,24-104-35,24-104-36,24-104-37, " & _
            "24-104-38,24-104-39,24-104-40,24-104-41,24-104-42,24-104-43, " & _
            "24-104-44, 24-104-45,24-104-46,24-104-47,24-104-48,24-104-49, " & _
            "24-104-50,24-104-51, 24-104-52,24-104-53,24-104-54,24-104-55, " & _
            "24-104-56,24-104-57,24-104-58, 24-104-59,24-104-60,24-104-61, " & _
            "24-104-62,24-104-63,24-104-64,24-104-65, 27-100-01,27-100-02, " & _
            "27-100-03,27-100-04,27-100-05,27-100-06,27-100-07, 27-101-01, " & _
            "27-102-01,27-102-02,27-103-01,27-104-01,27-105-01"
    
    If InStr(1, acuan2, kode, vbTextCompare) > 0 Then
        isKodeBuktiPotongValid = True
    Else
        isKodeBuktiPotongValid = False
    End If
End Function

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
    Dim File1 As String
    
    On Error GoTo er1
    MsgBox "Data di import per DIVISI!!", vbInformation
    File1 = "explorer.exe " & App.Path & "\rep\Template_ebupot26_baru.xls"
    'File1 = InputBox("", "", File1)
    Call Shell(File1, vbNormalFocus)
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION
End Sub

Private Sub Form_Load()
  Dim sql As String
  
  Me.Text1 = ""
  Me.txt_divisi = ""

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
      t1 = List1.text & Chr(13) & Chr(10)
      Call writefile(f, t1)
    Next
    Call closefile(f)
    MsgBox "File export di simpan di " & namaFile, vbInformation
    Me.Enable_Form
  End If
End Sub
