VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Ek_All_rekCab 
   ClientHeight    =   7245
   ClientLeft      =   240
   ClientTop       =   1050
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
   Begin VB.Frame Frame1 
      Caption         =   " Filter "
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   12015
      Begin VB.CommandButton cmd_load 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Load"
         Height          =   375
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cb_proyek 
         Height          =   330
         Left            =   1680
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   382
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Kode Proyek Lama"
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   435
         Width           =   1350
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   12015
      Begin VB.CommandButton cmd_ubah 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Ubah Data"
         Height          =   375
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton cmd_export 
         Caption         =   "&Export XLS"
         Height          =   375
         Left            =   10920
         TabIndex        =   6
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox txt_cari 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Text            =   "Text1"
         ToolTipText     =   "input dan ENTER"
         Top             =   4920
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4575
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   8070
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
         TabIndex        =   5
         Top             =   4995
         Width           =   705
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
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
      Caption         =   "Ekualisasi : Data Gabungan"
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
   Begin VB.Menu mnData 
      Caption         =   "Data"
      Begin VB.Menu mnSinkronisasi 
         Caption         =   "Sinkronisasi"
      End
      Begin VB.Menu mnHapusSemuaData 
         Caption         =   "Hapus Semua Data"
      End
   End
End
Attribute VB_Name = "frm_Ek_All_rekCab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim nama_data As String


Sub disable_Form()
    Me.Frame3.Enabled = False
    Me.Frame1.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame3.Enabled = True
    Me.Frame1.Enabled = True
End Sub


Function generate_sql() As String
    Dim sql As String, kondisi As String
    Dim cari As String, jmlData As Integer
    
    'kondisi
    kondisi = ""
    cari = ""
        
    sql = "Select id1, NO1, CABANG, " & _
        "DIVISI, NO_KONTRAK, NK_PPN, " & _
        "NAMA_OWNER, NAMA_PROYEK, KODE_PROYEK_LAMA, " & _
        "KODE_PROYEK_BARU, AKUN, DESKRIPSI, " & _
        "PU_2008, PU_2009, PU_2010, " & _
        "PU_2011, PU_2012, PU_2013,  " & _
        "PU_2014, PU_2015, PU_2016, " & _
        "PU_2017, PU_2018, PU_2019, " & _
        "PU_2020, PU_2021, PU_2022, " & _
        "Jml_PU, " & _
        "tgl_fp2008, no_fp2008, dpp_ppn2008,  " & _
        "tgl_fp2009, no_fp2009, dpp_ppn2009, " & _
        "tgl_fp2010, no_fp2010, dpp_ppn2010, " & _
        "tgl_fp2011, no_fp2011, dpp_ppn2011, " & _
        "tgl_fp2012, no_fp2012, dpp_ppn2012, " & _
        "tgl_fp2013, no_fp2013, dpp_ppn2013, " & _
        "tgl_fp2014, no_fp2014, dpp_ppn2014, " & _
        "tgl_fp2015, no_fp2015, dpp_ppn2015, " & _
        "tgl_fp2016, no_fp2016, dpp_ppn2016, " & _
        "tgl_fp2017, no_fp2017, dpp_ppn2017, " & _
        "tgl_fp2018, no_fp2018, dpp_ppn2018, " & _
        "tgl_fp2019, no_fp2019, dpp_ppn2019, " & _
        "tgl_fp2020, no_fp2020, dpp_ppn2020, " & _
        "tgl_fp2021, no_fp2021, dpp_ppn2021, " & _
        "tgl_fp2022, no_fp2022, dpp_ppn2022, "
    sql = sql & "total_dpp_ppn, SELISIH_pu_dppFP,  " & _
        "piutang11101, piutang11102, piutang11103, " & _
        "TOTAL_WIP_1, PRESTASI_YAMP, KARYA_YDF, " & _
        "TOTAL_WIP_2, TOTAL_WIP, selisih_dppfp_wip, " & _
        "PENJELASAN,  " & _
        "tgl_bp2008, no_bp2008, dpp_pph2008, " & _
        "tgl_bp2009, no_bp2009, dpp_pph2009, " & _
        "tgl_bp2010, no_bp2010, dpp_pph2010, " & _
        "tgl_bp2011, no_bp2011, dpp_pph2011, " & _
        "tgl_bp2012, no_bp2012, dpp_pph2012, " & _
        "tgl_bp2013, no_bp2013, dpp_pph2013, " & _
        "tgl_bp2014, no_bp2014, dpp_pph2014,  " & _
        "tgl_bp2015, no_bp2015, dpp_pph2015, " & _
        "tgl_bp2016, no_bp2016, dpp_pph2016, " & _
        "tgl_bp2017, no_bp2017, dpp_pph2017, " & _
        "tgl_bp2018, no_bp2018, dpp_pph2018, " & _
        "tgl_bp2019, no_bp2019, dpp_pph2019, " & _
        "tgl_bp2020, no_bp2020, dpp_pph2020, " & _
        "tgl_bp2021, no_bp2021, dpp_pph2021, " & _
        "tgl_bp2022, no_bp2022, dpp_pph2022,  " & _
        "total_dpp_PPh, dpp_fp_min_dpp_pph " & _
        "From all2016_all"
    
    
    If Not (Trim(Me.cb_proyek) = "" Or Trim(Me.cb_proyek) = "ALL") Then
        If Trim(kondisi) = "" Then
            kondisi = "kode_proyek_lama = '" & Trim(Me.cb_proyek) & "' "
        Else
            kondisi = kondisi & " and kode_proyek_lama = '" & Trim(Me.cb_proyek) & "' "
        End If
    End If
    
    
    
    '-- ini sql cari
    If Trim(Me.txt_cari.text) <> "" Then
        cari = "kode_proyek_baru like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NO_KONTRAK like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NAMA_OWNER like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NAMA_PROYEK like '%" & Trim(Me.txt_cari.text) & "%'"
    End If
    
    '-- gabungkan kondisi
    If Trim(kondisi) <> "" Then
        sql = sql & " where (" & kondisi & ") "
    End If
    
    '-- gabungkan cari
    If Trim(cari) <> "" Then
        If Trim(kondisi) <> "" Then
            sql = sql & " and (" & cari & ") "
        Else
            sql = sql & " where " & cari
        End If
    End If
        
    sql = sql & " order by kode_proyek_lama, kode_proyek_baru, id1"
        
    
    jmlData = cek_Int(InputBox("Jumlah data yang ditampilkan? (0:semua)", "", "0"))
    
    If jmlData > 0 Then
        sql = sql & " limit " & CStr(jmlData)
        Me.Frame3.Caption = " " & kondisi & " / " & Trim(Me.txt_cari.text) & " - Limit " & CStr(jmlData)
    Else
        Me.Frame3.Caption = " " & kondisi & " / " & Trim(Me.txt_cari.text)
    End If
    generate_sql = sql
End Function

Sub format_Grid()
    
    Dim jRec As Long
    Dim c As Integer
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
            
            'kecil
            If c = 0 Or c = 1 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).Width = 400
            End If
            
            'If c = 12 Or c = 20 Then
            '    Me.DataGrid1.Columns(c).Alignment = dbgCenter
            '    Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
            '    Me.DataGrid1.Columns(c).Width = 900
            'End If
    
            If c = 5 Or (c >= 12 And c <= 27) Or c = 30 Or c = 33 Or c = 36 Or c = 39 Or c = 42 Or _
                c = 45 Or c = 48 Or c = 51 Or c = 54 Or c = 57 Or c = 60 Or c = 63 Or c = 66 Or _
                c = 69 Or (c >= 72 And c <= 83) Or c = 87 Or c = 90 Or c = 93 Or c = 96 Or c = 99 Or _
                c = 102 Or c = 105 Or c = 108 Or c = 111 Or c = 114 Or c = 117 Or c = 120 Or c = 123 Or _
                c = 126 Or c = 129 Or c = 310 Or c = 131 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
End Sub



Private Sub cmd_export_Click()
    Dim jRec As Long
    
    Me.disable_Form
    jRec = RecordCount(rs)
    If jRec > 0 Then
        Call create_xls2(rs, "", "5, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, " & _
                        "27, 30, 33, 36, 39, 42, 45, 48, 51, 54, 57, 60, 63, 66, 69, 72, 73, 74, " & _
                        "75, 76, 77, 78, 79, 80, 81, 82, 83, 87, 90, 93, 96, 99, 102, 105, 108, " & _
                        "111, 114, 117, 120, 123, 126, 129, 130, 131,", "")
    End If
    Me.Enable_Form
End Sub


Private Sub LoadGrid()
    Dim sql As String, jRec As Long
    
    Me.disable_Form
    sql = generate_sql
    DoEvents
    
    If Trim(sql) <> "" Then
        Call dbMySQL_open
        If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
            sql = InputBox("", "", sql)
            MsgBox "error open sql"
            Me.Enable_Form
            Exit Sub
        End If
        
        Set Me.DataGrid1.DataSource = rs
        jRec = RecordCount(rs)
    End If
    Call format_Grid
    Me.Enable_Form
End Sub

Private Sub cmd_load_Click()
    Call LoadGrid
End Sub

Private Sub cmd_ubah_Click()
    Dim val1(11)
    Dim a As Integer
    Dim p
    Dim sql As String
    Dim start_kolom As Integer
    
    start_kolom = 12
    If RecordCount(rs) <= 0 Then Exit Sub
    
    p = MsgBox("Ubah data " & nama_data & "?", vbYesNo)
    If p = vbNo Then Exit Sub
        
        
    For a = start_kolom To rs.Fields.Count - 1
        val1(a) = InputBox(rs.Fields(a).Name, "Input", cek_null(rs.Fields(a).Value))
    Next
    
    sql = "update all2016_all set "
    For a = start_kolom To rs.Fields.Count - 1
        If a = rs.Fields.Count - 1 Then
            sql = sql & rs.Fields(a).Name & "='" & val1(a) & "'"
        Else
            sql = sql & rs.Fields(a).Name & "='" & val1(a) & "', "
        End If
    Next
    sql = sql & " where `id1` = '" & rs.Fields(0).Value & "'"
    
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("sql error", "", sql)
    Else
        Call LoadGrid
    End If
End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
  nama_data = "Bukti Potong"
  Call dbMySQL_open
    
  'load combo
  Me.txt_cari.text = ""
  sql = "select distinct kode_proyek_lama from all2016_all"
  Call Load_combo(Me.cb_proyek, sql, cnn, True, , 1)
  
  Me.Height = 8010
  Me.Width = 12420
  
  'get level
  '2 : operator gedung
  '3 : UKP
  
  '---------
  
  Level1 = tbPengguna_getLevel1(frMenu1.nmLogin)
  'If Level1 = 2 Then
  '  Me.cb_divisi.text = tbPengguna_getDivisi(frMenu1.nmLogin)
  '  Me.cb_divisi.Enabled = False
  'ElseIf Level1 = 3 Then
  'Else
  '  Call pesan2("Level tidak valid", , vbYellow)
 '   Me.cb_divisi.Enabled = False
 ' End If
 
 'Call LoadGrid
  Call pesan2("klik cari data dan ENTER")
End Sub


Private Sub Form_Resize()
    If Me.Width - 405 > 0 Then Me.Frame3.Width = Me.Width - 405
    If Me.Height - 2595 > 0 Then Me.Frame3.Height = Me.Height - 2595

    If Me.Width - 645 > 0 Then Me.DataGrid1.Width = Me.Width - 645
    If Me.Height - 3435 > 0 Then Me.DataGrid1.Height = Me.Height - 3435

    If Me.Height - 3090 > 0 Then Me.txt_cari.Top = Me.Height - 3090
    Me.Label6.Top = Me.txt_cari.Top
    Me.cmd_export.Top = Me.txt_cari.Top
    Me.cmd_ubah.Top = Me.txt_cari.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call dbMySQL_close
End Sub

Private Sub mnHapusSemuaData_Click()
    Dim p
    Dim sql As String
    
    p = MsgBox("Hapus semua data?", vbYesNo)
    If p = vbYes Then
        sql = "delete from all2016_all"
        If ExecSQL1(cnn, sql) <> 0 Then
                sql = InputBox("error", "", sql)
        End If
        Call LoadGrid
    End If
End Sub

Private Sub mnSinkronisasi_Click()
    frm_Ek_AllSin.Show
End Sub

Private Sub txt_cari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call LoadGrid
    End If
End Sub
