VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_EkMastProyek 
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
   Begin VB.Frame Frame3 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   12015
      Begin VB.CommandButton cmd_hapus 
         Caption         =   "Hapus Data (s)"
         Height          =   375
         Left            =   7200
         TabIndex        =   9
         ToolTipText     =   "Pilh data yang akan dihapus, dan klik "
         Top             =   5760
         Width           =   1215
      End
      Begin VB.CommandButton cmd_ubah 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Ubah Data"
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5753
         Width           =   1095
      End
      Begin VB.CommandButton cmd_add 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Tambah Data"
         Height          =   375
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5753
         Width           =   1095
      End
      Begin VB.CommandButton cmd_export 
         Caption         =   "&Export XLS"
         Height          =   375
         Left            =   10920
         TabIndex        =   6
         Top             =   5753
         Width           =   975
      End
      Begin VB.TextBox txt_cari 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Text            =   "Text1"
         ToolTipText     =   "input dan ENTER"
         Top             =   5753
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5415
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   9551
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
         Top             =   5835
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
      Caption         =   "Ekualisasi : Master Proyek"
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
   Begin VB.Menu mnMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnImport 
         Caption         =   "Import"
      End
      Begin VB.Menu mnHapusAll 
         Caption         =   "Hapus Semua Data"
      End
      Begin VB.Menu mnDataDouble 
         Caption         =   "Data Proyek Double"
      End
   End
End
Attribute VB_Name = "frm_EkMastProyek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset


Sub disable_Form()
    Me.Frame3.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame3.Enabled = True
End Sub


Function generate_sql() As String
    Dim sql As String, kondisi As String
    Dim cari As String
    
    'kondisi
    cari = ""
        
    sql = "Select ID, NO, CABANG, " & _
        "divisi, NO_KONTRAK, NK_PPN, " & _
        "OWNER, PROYEK, KODE_ACPAC, " & _
        "kode_Proyek_lama, kode_Proyek_baru, Description " & _
        "From all2016_master "
    
    '-- ini sql cari
    If Trim(Me.txt_cari.text) <> "" Then
        cari = "divisi like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NO_KONTRAK like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "OWNER like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "PROYEK like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "kode_Proyek_lama like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "kode_Proyek_baru like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "Description like '%" & Trim(Me.txt_cari.text) & "%' "
    End If
    
    '-- gabungkan cari
    If Trim(cari) <> "" Then
        sql = sql & " where " & cari
    End If
    
    generate_sql = sql & " order by kode_Proyek_lama"
    Me.Frame3.Caption = " " & kondisi & " / " & Trim(Me.txt_cari.text)
End Function

Sub format_Grid()
    
    Dim jRec As Long
    Dim c As Integer
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
    
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
            
            'kecil
            If c = 0 Or c = 1 Or c = 2 Or c = 3 Or c = 9 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).Width = 400
            End If
            
            'If c = 12 Or c = 20 Then
            '    Me.DataGrid1.Columns(c).Alignment = dbgCenter
            '    Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
            '    Me.DataGrid1.Columns(c).Width = 900
            'End If
    
            If c = 5 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
End Sub


Private Sub cmd_add_Click()
    Dim val1(11)
    Dim a As Integer
    Dim p
    Dim sql As String
    

    
    p = MsgBox("Tambah data Master Proyek Ekualisasi?", vbYesNo)
    If p = vbNo Then Exit Sub
        
    Call LoadGrid
        
    For a = 2 To rs.Fields.Count - 1
        val1(a) = InputBox(rs.Fields(a).Name, "Input", "")
    Next
    
    sql = "insert into all2016_master(" & _
            "CABANG, divisi, NO_KONTRAK, " & _
            "NK_PPN, OWNER, PROYEK, " & _
            "KODE_ACPAC, kode_Proyek_lama, kode_Proyek_baru, " & _
            "DESCRIPTION) values ("
    For a = 2 To rs.Fields.Count - 1
        If a = 2 Then
            sql = sql & "'" & val1(a) & "'"
        Else
            sql = sql & ",'" & val1(a) & "'"
        End If
    Next
    sql = sql & ")"
    
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("sql error", "", sql)
    Else
        Call LoadGrid
    End If
    
End Sub

Private Sub cmd_export_Click()
    Dim jRec As Long
    
    Me.disable_Form
    jRec = RecordCount(rs)
    If jRec > 0 Then
        Call create_xls2(rs, "", "", "")
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

Private Sub cmd_hapus_Click()
    Dim j As Integer, rec_no As Long
    Dim kode_proyek_lama As String, id1 As String, PROYEK As String
    Dim p
    Dim isAdaYangDihapus As Boolean, sql As String
    
    On Error GoTo er1
    isAdaYangDihapus = False
    For j = 0 To Me.DataGrid1.SelBookmarks.Count - 1
        rec_no = Me.DataGrid1.SelBookmarks.Item(j)
        rs.AbsolutePosition = rec_no
        
        id1 = cek_null(rs(0))
        kode_proyek_lama = cek_null(rs(9))
        PROYEK = cek_null(rs(7))
        p = MsgBox("Yakin menghapus 1 record data untuk " & vbCr & "Proyek: " & kode_proyek_lama & vbCr & _
                    "Nama : " & PROYEK & vbCr & "?", vbYesNo)
        If p = vbYes Then
            isAdaYangDihapus = True
            sql = "delete from all2016_master where ID = '" & id1 & "'"
            If ExecSQL1(cnn, sql) <> 0 Then
                sql = InputBox("error", "", sql)
            End If
        End If
    Next
    
    If isAdaYangDihapus = True Then Call LoadGrid
    
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
End Sub

Private Sub cmd_ubah_Click()
    Dim val1(11)
    Dim a As Integer
    Dim p
    Dim sql As String
    

    If RecordCount(rs) <= 0 Then Exit Sub
    
    p = MsgBox("Ubah data Master Proyek Ekualisasi?", vbYesNo)
    If p = vbNo Then Exit Sub
        
        
    For a = 2 To rs.Fields.Count - 1
        val1(a) = InputBox(rs.Fields(a).Name, "Input", rs.Fields(a).Value)
    Next
    
    sql = "update all2016_master set "
    For a = 2 To rs.Fields.Count - 1
        If a = rs.Fields.Count - 1 Then
            sql = sql & rs.Fields(a).Name & "='" & val1(a) & "'"
        Else
            sql = sql & rs.Fields(a).Name & "='" & val1(a) & "', "
        End If
    Next
    sql = sql & " where `id` = '" & rs.Fields(0).Value & "'"
    
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("sql error", "", sql)
    Else
        Call LoadGrid
    End If
End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
  Call dbMySQL_open
    
  'load combo
  Me.txt_cari.text = ""
  
  Me.Height = 7680
  Me.Width = 12390
  
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
    If Me.Height - 1455 > 0 Then Me.Frame3.Height = Me.Height - 1455

    
    If Me.Width - 645 > 0 Then Me.DataGrid1.Width = Me.Width - 645
    If Me.Height - 2295 > 0 Then Me.DataGrid1.Height = Me.Height - 2295

    If Me.Height - 1950 > 0 Then Me.txt_cari.Top = Me.Height - 1950
    Me.Label6.Top = Me.txt_cari.Top
    Me.cmd_export.Top = Me.txt_cari.Top
    Me.cmd_add.Top = Me.Label6.Top
    Me.cmd_ubah.Top = Me.txt_cari.Top
    Me.cmd_hapus.Top = Me.txt_cari.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call dbMySQL_close
End Sub

Private Sub mnDataDouble_Click()
    frm_Grid.Show
    frm_Grid.sql = "select kode_proyek_lama, kode_proyek_baru, count(*) as jumlah " & _
                    "From all2016_master " & _
                    "group by kode_proyek_lama, kode_proyek_baru having count(*) > 1"
    frm_Grid.judul = "Data Proyek Double"
    Call frm_Grid.LoadGrid(2, 2)
End Sub

Private Sub mnHapusAll_Click()
    Dim p
    Dim sql As String
    
    p = MsgBox("Hapus semua data?", vbYesNo)
    If p = vbYes Then
        sql = "delete from all2016_master"
        If ExecSQL1(cnn, sql) <> 0 Then
                sql = InputBox("error", "", sql)
        End If
        Call LoadGrid
    End If
End Sub

Private Sub mnImport_Click()
    frm_EkMastProyek_imp.Show
End Sub

Private Sub txt_cari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call LoadGrid
    End If
End Sub
