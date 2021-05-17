VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Ebupotpph26 
   ClientHeight    =   7260
   ClientLeft      =   165
   ClientTop       =   510
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
   ScaleHeight     =   7260
   ScaleWidth      =   12300
   Begin VB.Frame Frame1 
      Caption         =   " Filter "
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   12015
      Begin VB.ComboBox cb_masa 
         Height          =   330
         Left            =   2760
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cb_tahun 
         Height          =   330
         Left            =   840
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cb_divisi 
         Height          =   330
         Left            =   4800
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   360
         Width           =   1455
      End
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Divisi"
         Height          =   210
         Left            =   4320
         TabIndex        =   13
         Top             =   420
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Masa"
         Height          =   210
         Left            =   2280
         TabIndex        =   12
         Top             =   420
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun "
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   420
         Width           =   495
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
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton cmd_del 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Hapus Data"
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
      Top             =   7005
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
      Caption         =   "Data Upload eBupot 26"
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
      Begin VB.Menu mnImport 
         Caption         =   "Import"
      End
      Begin VB.Menu mnRekao 
         Caption         =   "Rekap"
      End
      Begin VB.Menu mnHapus 
         Caption         =   "Hapus"
      End
   End
End
Attribute VB_Name = "frm_Ebupotpph26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim nama_data As String


Sub update_label()
    Dim txt As String
    
    txt = "tahun " & Me.cb_tahun & " masa " & Me.cb_masa & " divisi " & Me.cb_divisi
    Me.mnHapus.Caption = "Hapus data " & txt
End Sub

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
        
    sql = "Select ID, NPWP_KPP, Kode_Proyek, " & _
        "No_Bukti_Akuntansi, Jenis_Dokumen, Tgl_Dokumen_, " & _
        "No_Faktur_Pajak, Kode_Form_Bukti_Potong, Masa_Pajak, " & _
        "Tahun_Pajak, Pembetulan, tgl_lahir_wp, " & _
        "TIN_, No_Paspor_WP_Terpotong, No_Kitas_WP_Terpotong, " & _
        "Nama_WP_yang_Dipotong, Alamat_WP_yang_Dipotong, Kode_Negara, " & _
        "Kode_Objek_Pajak, Penanda_tangan_BP_Pengurus, Nomor_Bukti_Potong, " & _
        "Tanggal_Bukti_Potong, Perkiraan_Penghasilan_Neto, Mendapatkan_Fasilitas, " & _
        "Nomor_Tanda_Terima_SKD, Tarif_SKD, Nomor_Aturan_DTP, NTPN_DTP, Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1,  " & _
        "Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, " & _
        "Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, " & _
        "Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, " & _
        "Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, " & _
        "Nilai_Bruto_6a_Nilai_Bruto_6, Tarif_6a_Tarif_6, PPh_Yang_Dipotong__6a_PPh_Yang_Dipotong__6, " & _
        "Nilai_Bruto_6b_Nilai_Bruto_7, Tarif_6b_Tarif_7, PPh_Yang_Dipotong__6b_PPh_Yang_Dipotong__7, " & _
        "Nilai_Bruto_6c_Nilai_Bruto_8, Tarif_6c_Tarif_8, PPh_Yang_Dipotong__6c_PPh_Yang_Dipotong__8, " & _
        "Nilai_Bruto_9, Tarif_9, PPh_Yang_Dipotong__9, " & _
        "Nilai_Bruto_10, Perkiraan_Penghasilan_Netto10,Tarif_10, " & _
        "PPh_Yang_Dipotong__10, Nilai_Bruto_11, Perkiraan_Penghasilan_Netto11, " & _
        "Tarif_11, PPh_Yang_Dipotong__11, Nilai_Bruto_12, " & _
        "Perkiraan_Penghasilan_Netto12, Tarif_12, PPh_Yang_Dipotong__12, " & _
        "Nilai_Bruto_13, Tarif_13, PPh_Yang_Dipotong__13, " & _
        "Kode_Jasa_6d1_PMK_244_PMK03_2008, Nilai_Bruto_6d1, Tarif_6d1, " & _
        "PPh_Yang_Dipotong__6d1, Kode_Jasa_6d2_PMK_244_PMK03_2008, Nilai_Bruto_6d2, " & _
        "Tarif_6d2, PPh_Yang_Dipotong__6d2, Kode_Jasa_6d3_PMK_244_PMK03_2008, "
    sql = sql & "Nilai_Bruto_6d3, Tarif_6d3, PPh_Yang_Dipotong__6d3, " & _
        "Kode_Jasa_6d4_PMK_244_PMK03_2008, Nilai_Bruto_6d4, Tarif_6d4, " & _
        "PPh_Yang_Dipotong__6d4, Kode_Jasa_6d5_PMK_244_PMK03_2008, Nilai_Bruto_6d5, " & _
        "Tarif_6d5, PPh_Yang_Dipotong__6d5, Kode_Jasa_6d6_PMK_244_PMK03_2008, " & _
        "Nilai_Bruto_6d6, Tarif_6d6, PPh_Yang_Dipotong__6d6, " & _
        "Jumlah_Nilai_Bruto_, Jumlah_PPh_Yang_Dipotong, tgl_import, " & _
        "kode_divisi " & _
        "From ebupot26 "
    
    If Not (Trim(Me.cb_tahun) = "" Or Trim(Me.cb_tahun) = "ALL") Then
        kondisi = "Tahun_Pajak = '" & Trim(Me.cb_tahun) & "' "
    End If
    
    If Not (Trim(Me.cb_masa) = "" Or Trim(Me.cb_masa) = "ALL") Then
        If Trim(kondisi) = "" Then
            kondisi = "Masa_Pajak = '" & Trim(Me.cb_masa) & "' "
        Else
            kondisi = kondisi & " and Masa_Pajak = '" & Trim(Me.cb_masa) & "' "
        End If
    End If
    
    If Not (Trim(Me.cb_divisi) = "" Or Trim(Me.cb_divisi) = "ALL") Then
        If Trim(kondisi) = "" Then
            kondisi = "kode_divisi = '" & Trim(Me.cb_divisi) & "' "
        Else
            kondisi = kondisi & " and kode_divisi = '" & Trim(Me.cb_divisi) & "' "
        End If
    End If
    
    
    
    '-- ini sql cari
    If Trim(Me.txt_cari.text) <> "" Then
        cari = "NPWP_KPP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "No_Bukti_Akuntansi like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "No_Faktur_Pajak like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "Tgl_Dokumen_ like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "tgl_lahir_wp like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "No_Paspor_WP_Terpotong like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "No_Kitas_WP_Terpotong like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "TIN_ like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "Nama_WP_yang_Dipotong like '%" & Trim(Me.txt_cari.text) & "%' "
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
        
    sql = sql & " order by Tahun_Pajak desc, Masa_Pajak, Kode_Proyek, Tgl_Dokumen_ "
        
    If cari = "" Then
        jmlData = cek_Int(InputBox("Jumlah data yang ditampilkan? (0:semua)", "", "0"))
        If jmlData > 0 Then
            sql = sql & " limit " & CStr(jmlData)
        End If
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
            If c = 0 Or c = 1 Or c = 2 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).Width = 400
            End If
            
            'If c = 12 Or c = 20 Then
            '    Me.DataGrid1.Columns(c).Alignment = dbgCenter
            '    Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
            '    Me.DataGrid1.Columns(c).Width = 900
            'End If
    
            If c = 7 Or c = 8 Then
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
    Dim sql As String, start_kolom As Integer
    

    start_kolom = 1
    p = MsgBox("Tambah data " & nama_data & "?", vbYesNo)
    If p = vbNo Then Exit Sub
        
    Call LoadGrid
        
    For a = start_kolom To rs.Fields.Count - 1
        val1(a) = InputBox(rs.Fields(a).Name, "Input", "")
    Next
    
    sql = "insert into all2016_fp(" & _
            "kode_divisi, kode_proyek_lama, kode_proyek_baru, " & _
            "tahun, tgl_fp, no_fp, " & _
            "dpp, ppn, keterangan) values ("
    For a = start_kolom To rs.Fields.Count - 1
        If a = start_kolom Then
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

Private Sub cb_divisi_Click()
    Call update_label
End Sub

Private Sub cb_masa_Click()
    Call update_label
End Sub

Private Sub cb_tahun_Click()
    Call update_label
End Sub

Private Sub cmd_del_Click()
    Dim id As String, p
    Dim nmTabel As String
    Dim klm(), isi()
    
    nmTabel = "ebupot26"
    klm = Array("ID")
    
    
    If RecordCount(rs) < 0 Then
        Call pesan2("tidak ada data")
        Exit Sub
    End If
    
    id = cek_null(rs(0))
    p = MsgBox("data id " & id & " yakin dihapus ? ", vbYesNo)
    If p = vbYes Then
        isi = Array(id)
        If tbDelete(nmTabel, klm, isi, cnn) = True Then
            Call pesan2("delete sukses")
            Call LoadGrid
        End If
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
        'sql = InputBox("", "", sql)
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
    Dim val1(100)
    Dim a As Integer
    Dim p
    Dim sql As String
    Dim start_kolom As Integer
    
    start_kolom = 5
    If RecordCount(rs) <= 0 Then Exit Sub
    
    p = MsgBox("Ubah data " & nama_data & "?", vbYesNo)
    If p = vbNo Then Exit Sub
        
        
    For a = start_kolom To 95
        If a <= 27 Or a >= 94 Then
            val1(a) = InputBox(rs.Fields(a).Name, "Input", cek_null(rs.Fields(a).Value))
        End If
    Next
    
    p = MsgBox("Ubah data " & nama_data & "?", vbYesNo)
    If p = vbNo Then Exit Sub
    
    sql = "update ebupot26 set "
    For a = start_kolom To 95
        If a <= 27 Or a >= 94 Then
            If a = 95 Then
                sql = sql & rs.Fields(a).Name & "='" & val1(a) & "'"
            Else
                sql = sql & rs.Fields(a).Name & "='" & val1(a) & "', "
            End If
        End If
    Next
    sql = sql & " where `id` = '" & rs.Fields(0).Value & "'"
    
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("sql error", "", sql)
    Else
        Call LoadGrid
    End If
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
  nama_data = "ebupot26"
  Call dbMySQL_open
    
  'load combo
  Me.txt_cari.text = ""
  
  sql = "select distinct Tahun_Pajak from ebupot26"
  Call Load_combo(Me.cb_tahun, sql, cnn, True, , 1)
  
  sql = "select distinct Masa_Pajak from ebupot26"
  Call Load_combo(Me.cb_masa, sql, cnn, True, , 1)
  
  sql = "select distinct kode_divisi from ebupot26"
  Call Load_combo(Me.cb_divisi, sql, cnn, True, , 1)
  
  Me.Height = 8010
  Me.Width = 12420
  
  'get level
  '2 : operator gedung
  '3 : UKP
  
  '---------
  
  Level1 = tbPengguna_getLevel1(frMenu1.nmLogin)
  If Level1 = 2 Then
    Me.cb_divisi.text = tbPengguna_getDivisi(frMenu1.nmLogin)
    Me.cb_divisi.Enabled = False
  ElseIf Level1 = 3 Then
  Else
    Call pesan2("Level tidak valid", , vbYellow)
   Me.cb_divisi.Enabled = False
  End If
 
 'Call LoadGrid
  Call pesan2("Pilih Filter dan klik 'LOAD', atau " & vbCr & _
                "klik cari data dan ENTER")
End Sub


Private Sub Form_Resize()
    If Me.Width - 405 > 0 Then Me.Frame3.Width = Me.Width - 405
    If Me.Height - 2595 > 0 Then Me.Frame3.Height = Me.Height - 2595

    If Me.Width - 645 > 0 Then Me.DataGrid1.Width = Me.Width - 645
    If Me.Height - 3435 > 0 Then Me.DataGrid1.Height = Me.Height - 3435

    If Me.Height - 3090 > 0 Then Me.txt_cari.Top = Me.Height - 3090
    Me.Label6.Top = Me.txt_cari.Top
    Me.cmd_export.Top = Me.txt_cari.Top
    Me.cmd_del.Top = Me.Label6.Top
    Me.cmd_ubah.Top = Me.txt_cari.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call dbMySQL_close
End Sub

Private Sub mnHapus_Click()
Dim p, sql As String, kondisi As String
    
    p = MsgBox(mnHapus.Caption & "?", vbYesNo)
    If p = vbNo Then
        Call pesan2("dibatalkan")
        Exit Sub
    End If
    
    '--- prepare
    sql = "delete from ebupot26"
    If Not (Trim(Me.cb_tahun) = "" Or Trim(Me.cb_tahun) = "ALL") Then
        kondisi = "Tahun_Pajak = '" & Trim(Me.cb_tahun) & "' "
    End If
    
    If Not (Trim(Me.cb_masa) = "" Or Trim(Me.cb_masa) = "ALL") Then
        If Trim(kondisi) = "" Then
            kondisi = "Masa_Pajak = '" & Trim(Me.cb_masa) & "' "
        Else
            kondisi = kondisi & " and Masa_Pajak = '" & Trim(Me.cb_masa) & "' "
        End If
    End If
    
    If Not (Trim(Me.cb_divisi) = "" Or Trim(Me.cb_divisi) = "ALL") Then
        If Trim(kondisi) = "" Then
            kondisi = "kode_divisi = '" & Trim(Me.cb_divisi) & "' "
        Else
            kondisi = kondisi & " and kode_divisi = '" & Trim(Me.cb_divisi) & "' "
        End If
    End If
    
        
    '-- gabungkan kondisi
    If Trim(kondisi) <> "" Then
        sql = sql & " where (" & kondisi & ") "
    End If
    
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("error", "", sql)
    Else
        Call pesan2("data di hapus")
        Call LoadGrid
    End If
End Sub

Private Sub mnImport_Click()
    frm_Ebupotpph26_imp.Show
End Sub


Private Sub mnRekao_Click()
    Dim sql As String
    Dim kondisi As String
    
    If Trim(Me.cb_tahun) = "ALL" Or Trim(Me.cb_tahun) = "" Then
        kondisi = ""
    Else
        kondisi = "Tahun_Pajak = '" & Trim(Me.cb_tahun) & "' "
    End If
    
    If Trim(Me.cb_divisi) = "ALL" Or Trim(Me.cb_divisi) = "" Then
    Else
        If Trim(kondisi) = "" Then
            kondisi = "kode_divisi = '" & Trim(Me.cb_divisi) & "' "
        Else
            kondisi = kondisi & " and kode_divisi = '" & Trim(Me.cb_divisi) & "' "
        End If
    End If
    
    
    frm_Grid.Show
    sql = "select Tahun_Pajak, kode_divisi, sum(Jumlah_Nilai_Bruto_), " & _
            "sum(Jumlah_PPh_Yang_Dipotong) " & _
            "from ebupot26 "
    If Trim(kondisi) <> "" Then
        sql = sql & "where " & kondisi
    End If
    sql = sql & "group by Tahun_Pajak, kode_divisi"
    
    frm_Grid.sql = sql
    frm_Grid.judul = "Rekap " & kondisi
    Call frm_Grid.LoadGrid(2, 3)
End Sub

Private Sub txt_cari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call LoadGrid
    End If
End Sub
