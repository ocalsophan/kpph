VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_mKaryawan 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7245
   ClientLeft      =   225
   ClientTop       =   1035
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
      TabIndex        =   10
      Top             =   2040
      Width           =   12015
      Begin VB.CommandButton cmd_export 
         Caption         =   "&Export XLS"
         Height          =   375
         Left            =   10920
         TabIndex        =   15
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmd_hapus1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Hapus Data(s)"
         Height          =   375
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CommandButton cmd_edit 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Edit Data"
         Height          =   375
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4320
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3975
         Left            =   120
         TabIndex        =   5
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "  Filter "
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
      Width           =   12015
      Begin VB.TextBox txt_Npwp 
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txt_Alamat 
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txt_Nama 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txt_Nik 
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "NPWP"
         Height          =   210
         Left            =   3600
         TabIndex        =   12
         Top             =   915
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Alamat"
         Height          =   210
         Left            =   3600
         TabIndex        =   11
         Top             =   435
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nama"
         Height          =   210
         Left            =   360
         TabIndex        =   9
         Top             =   915
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NIK"
         Height          =   210
         Left            =   360
         TabIndex        =   8
         Top             =   435
         Width           =   240
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
      Caption         =   "Master Karyawan"
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
   Begin VB.Menu mnFile 
      Caption         =   "File"
      Begin VB.Menu mnTambah 
         Caption         =   "Tambah Data"
      End
      Begin VB.Menu mnImport 
         Caption         =   "Import Data Karyawan"
      End
   End
End
Attribute VB_Name = "frm_mKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset

Sub disable_Form()
    Me.Frame2.Enabled = False
    Me.Frame3.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame2.Enabled = True
    Me.Frame3.Enabled = True
End Sub


Function generate_sql() As String
    Dim sql As String, cari As String
    Dim limit1 As String
    
    '-- referensi
    '0: id1, nik, npwp, " & _
    '3: "nama, alamat, jenis_kelamin, " & _
    '6: "ptkp
    '-----------
    
    sql = "select id1, nik, npwp, " & _
            "nama, alamat, jenis_kelamin, " & _
            "ptkp, '' as status " & _
            "from mkaryawan "
    
    '-- fiter cari
    cari = ""
        
    If Trim(Me.txt_Nik) <> "" Then
        If Trim(cari) <> "" Then cari = cari & " and "
        cari = cari & "nik like '%" & Trim(Me.txt_Nik.Text) & "%' "
    End If
    
    If Trim(Me.txt_Nama) <> "" Then
        If Trim(cari) <> "" Then cari = cari & " and "
        cari = cari & "nama like '%" & Trim(Me.txt_Nama.Text) & "%' "
    End If
    
    If Trim(Me.txt_Alamat) <> "" Then
        If Trim(cari) <> "" Then cari = cari & " and "
        cari = cari & "alamat like '%" & Trim(Me.txt_Alamat.Text) & "%' "
    End If
    
    If Trim(Me.txt_Npwp) <> "" Then
        If Trim(cari) <> "" Then cari = cari & " and "
        cari = cari & "npwp like '%" & Trim(Me.txt_Npwp.Text) & "%' "
    End If
    '-----------------
    
    If Trim(cari) <> "" Then
        sql = sql & "where " & cari
    End If
    
    If Trim(cari) = "" Then
        limit1 = cek_Int(InputBox("Limit", "", "5000"), 5000)
        generate_sql = sql & " order by nama, npwp, nik, alamat, jenis_kelamin limit " & limit1
    Else
        generate_sql = sql & " order by nama, npwp, nik, alamat, jenis_kelamin "
    End If
    
    'sql = InputBox("", "", sql)
End Function

Sub format_Grid()
    
    Dim jenisPPh As String
    Dim jRec As Long
    Dim c As Integer
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
    
    '-- referensi
    '0: id1, nik, npwp, " & _
    '3: "nama, alamat, jenis_kelamin, " & _
    '6: "ptkp
    '-----------
    
        
        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
            
            'kecil
            If c = 0 Or c = 5 Or c = 6 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).Width = 700
            End If
    
            'If c = 23 Then
            '    Me.DataGrid1.Columns(c).Alignment = dbgCenter
            '    Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
            '    Me.DataGrid1.Columns(c).Width = 900
            'End If
    
            'If c = 13 Or c = 14 Or c = 15 Or c = 16 Or c = 17 Or c = 18 Or c = 19 Or _
            '    c = 20 Or c = 21 Or c = 24 Or c = 25 Then
                
            '    Me.DataGrid1.Columns(c).Alignment = dbgRight
            '    Me.DataGrid1.Columns(c).NumberFormat = "###,###"
            '    Me.DataGrid1.Columns(c).Width = 1400
            'End If
        Next

End Sub


Private Sub cmd_edit_Click()
    '-- referensi
    '0: id1, nik, npwp, " & _
    '3: "nama, alamat, jenis_kelamin, " & _
    '6: "ptkp
    '-----------
    
    Dim id1 As String, NIK As String, npwp As String
    Dim nama As String, alamat As String, jenis_kelamin As String
    Dim ptkp As String, p
    Dim klm(), isi()
    
    id1 = cek_null(rs(0))
    NIK = cek_null(rs(1))
    npwp = cek_null(rs(2))
    nama = cek_null(rs(3))
    alamat = cek_null(rs(4))
    jenis_kelamin = cek_null(rs(5))
    ptkp = cek_null(rs(6))
    
    If Trim(id1) = "" Then Exit Sub
    p = MsgBox("Ubah data " & nama & " / " & NIK & "?", vbYesNo)
    If p = vbYes Then
        nama = InputBox("Input Nama", "", nama)
        npwp = cleanNpwp(InputBox("Input NPWP", "", npwp))
        NIK = InputBox("Input NIK", "", NIK)
        alamat = InputBox("Input Alamat", "", alamat)
        jenis_kelamin = InputBox("Input Jenis Kelamin", "", jenis_kelamin)
        ptkp = InputBox("Input PTKP", "", ptkp)
    
        If Trim(nama) = "" Or Trim(NIK) = "" Or Trim(npwp) = "" Or Trim(alamat) = "" Or Trim(jenis_kelamin) = "" _
            Or Trim(ptkp) = "" Then
            
            Call pesan2("ada data yang kosong. Update di batalkan", , vbYellow)
            Exit Sub
        End If
    
        'update
        klm = Array("nik", "npwp", "nama", "alamat", "jenis_kelamin", "ptkp")
        isi = Array(NIK, npwp, nama, alamat, jenis_kelamin, ptkp)
    
        If tbUpdate("mkaryawan", klm, isi, cnn, "id1 = " & id1) = True Then
            Call load_grid
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
    
    If Trim(Me.txt_Nik) <> "" Then
        If Trim(judul) <> "" Then judul = judul & " and "
        judul = judul & "Filter NIK " & Me.txt_Nik
    End If
    
    If Trim(Me.txt_Nama) <> "" Then
        If Trim(judul) <> "" Then judul = judul & " and "
        judul = judul & "Filter Nama " & Me.txt_Nama
    End If
    
    If Trim(Me.txt_Alamat) <> "" Then
        If Trim(judul) <> "" Then judul = judul & " and "
        judul = judul & "Filter Alamat " & Me.txt_Alamat
    End If
    
    If Trim(Me.txt_Npwp) <> "" Then
        If Trim(judul) <> "" Then judul = judul & " and "
        judul = judul & "Filter NPWP " & Me.txt_Npwp
    End If
    
    If Trim(judul) <> "" Then
        judul = judul
    Else
        judul = "no Filter"
    End If
    '-------------
    
    Me.disable_Form
    
    '-- referensi
    '0: id1, nik, npwp, " & _
    '3: "nama, alamat, jenis_kelamin, " & _
    '6: "ptkp
    '-----------
    
    jRec = RecordCount(rs)
    If jRec > 0 Then
        Call create_xls2(rs, judul, "", "", "", "01")
    End If
    Me.Enable_Form
End Sub


Private Sub cmd_hapus1_Click()
    Dim j As Integer, rec_no As Long
    Dim nama As String, id1 As String
    Dim p
    Dim isAdaYangDihapus As Boolean
    Dim klm(), isi()
    
    '-- referensi
    '0: id1, nik, npwp, " & _
    '3: "nama, alamat, jenis_kelamin, " & _
    '6: "ptkp
    '-----------
    
    On Error GoTo er1
    isAdaYangDihapus = False
    
    If Me.DataGrid1.SelBookmarks.Count <= 0 Then
        Call pesan2("tidak ada data yang di pilih", , vbYellow)
    End If
    
    For j = 0 To Me.DataGrid1.SelBookmarks.Count - 1
        rec_no = Me.DataGrid1.SelBookmarks.Item(j)
        rs.AbsolutePosition = rec_no
        
        nama = cek_null(rs(3))
        id1 = cek_null(rs(0))
        p = MsgBox("Yakin menghapus 1 record data untuk " & vbCr & "Nama: " & nama & vbCr & _
                    "ID : " & id1 & vbCr & "?", vbYesNo)
        If p = vbYes Then
            klm = Array("id1")
            isi = Array(id1)
            Call tbDelete("mkaryawan", klm, isi, cnn)
            isAdaYangDihapus = True
        End If
    Next
    
    If isAdaYangDihapus = True Then Call load_grid
    
    Exit Sub
er1:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub load_grid()
    Dim sql As String, npwp As String
    Dim rsLoad As ADODB.Recordset
    Dim a As Integer, c As Long, jRec As Long
    
    Me.disable_Form
    sql = generate_sql
    DoEvents
    'MsgBox sql
    If Trim(sql) <> "" Then
        Call dbMySQL_open
        If OpenRecordSet(cnn, rsLoad, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
            sql = InputBox("sql error", "", sql)
            Me.Enable_Form
            Exit Sub
        End If
        
        'create copy rs
        If createRS_duplicate(rsLoad, rs) = True Then
            jRec = RecordCount(rsLoad)
            If jRec > 0 Then
                rsLoad.MoveFirst
                c = 1
                Do While rsLoad.EOF = False
                    Call info_progress(Me.StatusBar1, 1, c, jRec, "load karyawan")
                    rs.AddNew
                    For a = 0 To rsLoad.Fields.Count - 1
                        If a = 7 Then
                            'If npwp = "818797870502000" Then
                            '    MsgBox "a"
                            'End If
                            npwp = cek_null(rsLoad(2))
                            If checkNPWP(npwp) = True Then
                                rs.Fields(a).Value = "-"
                            Else
                                rs.Fields(a).Value = "NPWP notValid"
                            End If
                        Else
                            rs.Fields(a).Value = rsLoad.Fields(a).Value
                        End If
                    Next
                    rs.Update
                    c = c + 1
                    rsLoad.MoveNext
                Loop
            End If
            
            
            
            Set Me.DataGrid1.DataSource = rs
            Call format_Grid
            Call info(1, "Jumlah data=" & RecordCount(rs), Me.StatusBar1)
        End If
        
        
    End If
    Me.Enable_Form
End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
    
  Me.txt_Nik.Text = ""
  Me.txt_Nama.Text = ""
  Me.txt_Alamat.Text = ""
  Me.txt_Npwp.Text = ""
  
  '---------------------
    If dbMySQL_open = True Then
    Else
        MsgBox "Open Database Tidak Sukses", vbExclamation
    End If
    '---------------------
    Call load_grid
End Sub


Private Sub mnImport_Click()
    frm_impMKaryawan.Show
End Sub

Private Sub mnTambah_Click()
    Dim NIK As String, npwp As String
    Dim nama As String, alamat As String, jenis_kelamin As String
    Dim ptkp As String, p
    Dim klm(), isi()
    
    
    p = MsgBox("Tambah data Karyawan ?", vbYesNo)
    If p = vbYes Then
        nama = InputBox("Input Nama", "", nama)
        npwp = cleanNpwp(InputBox("Input NPWP", "", npwp))
        NIK = InputBox("Input NIK", "", NIK)
        alamat = InputBox("Input Alamat", "", alamat)
        jenis_kelamin = InputBox("Input Jenis Kelamin", "", jenis_kelamin)
        If jenis_kelamin = "P" Or jenis_kelamin = "L" Then
        Else
            jenis_kelamin = "P"
        End If
        
        ptkp = InputBox("Input PTKP", "", ptkp)
        If tbM_Ptkp_isDataAda(ptkp) = True Then
        Else
            ptkp = "TK"
        End If
    
        If Trim(nama) = "" Or Trim(NIK) = "" Or Trim(npwp) = "" Or Trim(alamat) = "" Or Trim(jenis_kelamin) = "" _
            Or Trim(ptkp) = "" Then
            
            Call pesan2("ada data yang kosong. Insert di batalkan", , vbYellow)
            Exit Sub
        End If
    
        'update
        klm = Array("nik", "npwp", "nama", "alamat", "jenis_kelamin", "ptkp")
        isi = Array(NIK, npwp, nama, alamat, jenis_kelamin, ptkp)
    
        If tbInsert("mkaryawan", klm, isi, cnn) = True Then
            Call load_grid
        Else
            Call pesan2("error edit", 5, vbYellow)
        End If
    Else
        Call pesan2("Batal", 5, vbYellow)
    End If
End Sub

Private Sub txt_Alamat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call load_grid
End Sub

Private Sub txt_Nama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call load_grid
End Sub

Private Sub txt_Nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call load_grid
End Sub

Private Sub txt_Npwp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call load_grid
End Sub
