VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_EkPpnAll 
   ClientHeight    =   7245
   ClientLeft      =   240
   ClientTop       =   750
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
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   12015
      Begin VB.CommandButton cmd_export 
         Caption         =   "&Export XLS"
         Height          =   375
         Left            =   10920
         TabIndex        =   11
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txt_cari 
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Text            =   "Text1"
         ToolTipText     =   "input dan ENTER"
         Top             =   4320
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3975
         Left            =   120
         TabIndex        =   8
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
         TabIndex        =   10
         Top             =   4402
         Width           =   705
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
      TabIndex        =   4
      Top             =   600
      Width           =   12015
      Begin VB.CommandButton cmd_load 
         BackColor       =   &H00C0FFC0&
         Caption         =   "2. Load"
         Height          =   375
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox cb_proyek 
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
         Caption         =   "Proyek"
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cab / Div"
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   420
         Width           =   645
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
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
      Caption         =   "Data Ekualisasi PPN"
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
Attribute VB_Name = "frm_EkPpnAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset





Sub disable_Form()
    Me.Frame1.Enabled = False
    Me.Frame3.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame1.Enabled = True
    Me.Frame3.Enabled = True
End Sub


Function generate_sql() As String
    Dim sql As String, kondisi As String
    Dim cari As String
    
    'kondisi
    kondisi = ""
    cari = ""
        
    sql = "Select all2016.id, all2016.no, CABANG_DIVISI, " & _
        "NO_KONTRAK, NK_PPN, OWNER, " & _
        "PROYEK, KODE_ACPAC, kode_Proyek, " & _
        "DESCRIPTION, PU_2008, PU_2009, " & _
        "PU_2010, PU_2011, PU_2012, " & _
        "PU_2013, PU_2014, PU_2015, " & _
        "PU_2016, PU_2017, PU_2018, " & _
        "PU_2019, PU_2020, Jumlah, " & _
        "NOFP_2008, DPP_2008, NOFP_2009,  " & _
        "DPP_2009, NOFP_2010, DPP_2010, " & _
        "NOFP_2011, DPP_2011, NOFP_2012, " & _
        "DPP_2012, NOFP_2013, DPP_2013, " & _
        "NOFP_2014, DPP_2014, NOFP_2015,  " & _
        "DPP_2015, NOFP_2016, DPP_2016,  " & _
        "NOFP_2017, DPP_2017, NOFP_2018, " & _
        "DPP_2018, NOFP_2019, DPP_2019,  " & _
        "NOFP_2020, DPP_2020, total_dpp_all, " & _
        "SELISIH, PENJELASAN  " & _
        "From all2016 "
    
    '-- ini sql kondisi
    If Trim(Me.cb_divisi.text) = "ALL" Or Trim(Me.cb_divisi.text) = "" Then
    Else
        kondisi = "CABANG_DIVISI = '" & Trim(Me.cb_divisi.text) & "'"
    End If
    
    If Trim(Me.cb_proyek.text) = "ALL" Or Trim(Me.cb_proyek.text) = "" Then
    Else
        If Trim(kondisi) = "" Then
            kondisi = "kode_Proyek = '" & Trim(Me.cb_proyek.text) & "'"
        Else
            kondisi = kondisi & "and kode_Proyek = '" & Trim(Me.cb_proyek.text) & "'"
        End If
        
        
    End If
    
    
    '-- ini sql cari
    If Trim(Me.txt_cari.text) <> "" Then
        cari = "OWNER like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "DESCRIPTION like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NOFP_2008 like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NOFP_2009 like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NOFP_2010 like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NOFP_2011 like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NOFP_2012 like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NOFP_2013 like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NOFP_2014 like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NOFP_2015 like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NOFP_2016 like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NOFP_2017 like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NOFP_2018 like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NOFP_2019 like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                "NOFP_2020 like '%" & Trim(Me.txt_cari.text) & "%' "
    End If
    
    '-- gabungkan kondisi
    If Trim(kondisi) <> "" Then
        sql = sql & " where (" & kondisi & ") "
    End If
    
    '-- gabungkan cari
    If Trim(cari) <> "" Then
        If Trim(kondisi) = "" Then
            sql = sql & " where " & cari
        Else
            sql = sql & " and (" & cari & ") "
        End If
    End If
    
    generate_sql = sql & " order by `ID`"
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
            If c = 0 Or c = 1 Or c = 2 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).Width = 400
            End If
            
            'If c = 12 Or c = 20 Then
            '    Me.DataGrid1.Columns(c).Alignment = dbgCenter
            '    Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
            '    Me.DataGrid1.Columns(c).Width = 900
            'End If
    
            If c = 4 Or c = 10 Or c = 11 Or c = 12 Or c = 13 Or c = 14 Or _
                c = 15 Or c = 16 Or c = 17 Or c = 18 Or c = 19 Or c = 20 _
                Or c = 21 Or c = 22 Or c = 23 Or c = 25 Or c = 29 Or c = 29 _
                Or c = 31 Or c = 33 Or c = 35 Or c = 37 Or c = 39 Or c = 41 _
                Or c = 43 Or c = 45 Or c = 47 Or c = 49 Or c = 50 Or c = 51 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
End Sub

Private Sub cb_divisi_Click()
    Dim kdDivisi As String
    
    kdDivisi = Trim(Me.cb_divisi.text)
    If kdDivisi = "ALL" Then kdDivisi = ""

    Call tbAll2016_loadProyek(Me.cb_proyek, kdDivisi)
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


Private Sub cmd_Load_Click()
    Dim sql As String, jRec As Long
    Dim PU_2008 As Currency, PU_2009 As Currency, PU_2010 As Currency
    Dim PU_2011 As Currency, PU_2012 As Currency, PU_2013 As Currency
    Dim PU_2014 As Currency, PU_2015 As Currency, PU_2016 As Currency
    Dim PU_2017 As Currency, PU_2018 As Currency, PU_2019 As Currency
    Dim PU_2020 As Currency, Jumlah As Currency
    Dim dpp_2008 As Currency, dpp_2009 As Currency, dpp_2010 As Currency
    Dim dpp_2011 As Currency, dpp_2012 As Currency, dpp_2013 As Currency
    Dim dpp_2014 As Currency, dpp_2015 As Currency, dpp_2016 As Currency
    Dim dpp_2017 As Currency, dpp_2018 As Currency, dpp_2019 As Currency
    Dim dpp_2020 As Currency, total_dpp_all As Currency
    
    Me.disable_Form
    sql = generate_sql
    DoEvents
    
    If Trim(sql) <> "" Then
        Call dbMySQL_open
        If OpenRecordSet(cnn, rs, sql, adOpenDynamic, adLockPessimistic, adUseClient) <> 0 Then
            sql = InputBox("", "", sql)
            MsgBox "error open sql"
            Me.Enable_Form
            Exit Sub
        End If
        
        Set Me.DataGrid1.DataSource = rs
        jRec = RecordCount(rs)
        
        '-- hitung
        PU_2008 = 0
        PU_2009 = 0
        PU_2010 = 0
        PU_2011 = 0
        PU_2012 = 0
        PU_2013 = 0
        PU_2014 = 0
        PU_2015 = 0
        PU_2016 = 0
        PU_2017 = 0
        PU_2018 = 0
        PU_2019 = 0
        PU_2020 = 0
        Jumlah = 0
        dpp_2008 = 0
        dpp_2009 = 0
        dpp_2010 = 0
        dpp_2011 = 0
        dpp_2012 = 0
        dpp_2013 = 0
        dpp_2014 = 0
        dpp_2015 = 0
        dpp_2016 = 0
        dpp_2017 = 0
        dpp_2018 = 0
        dpp_2019 = 0
        dpp_2020 = 0
        total_dpp_all = 0
        
        rs.MoveFirst
        Do While rs.EOF = False
            PU_2008 = PU_2008 + cek_Money(rs(10))
            PU_2009 = PU_2009 + cek_Money(rs(11))
            PU_2010 = PU_2010 + cek_Money(rs(12))
            PU_2011 = PU_2011 + cek_Money(rs(13))
            PU_2012 = PU_2012 + cek_Money(rs(14))
            PU_2013 = PU_2013 + cek_Money(rs(15))
            PU_2014 = PU_2014 + cek_Money(rs(16))
            PU_2015 = PU_2015 + cek_Money(rs(17))
            PU_2016 = PU_2016 + cek_Money(rs(18))
            PU_2017 = PU_2017 + cek_Money(rs(19))
            PU_2018 = PU_2018 + cek_Money(rs(20))
            PU_2019 = PU_2019 + cek_Money(rs(21))
            PU_2020 = PU_2020 + cek_Money(rs(22))
            Jumlah = Jumlah + cek_Money(rs(23))
            
            dpp_2008 = dpp_2008 + cek_Money(rs(25))
            dpp_2009 = dpp_2009 + cek_Money(rs(27))
            dpp_2010 = dpp_2010 + cek_Money(rs(29))
            dpp_2011 = dpp_2011 + cek_Money(rs(31))
            dpp_2012 = dpp_2012 + cek_Money(rs(33))
            dpp_2013 = dpp_2013 + cek_Money(rs(35))
            dpp_2014 = dpp_2014 + cek_Money(rs(37))
            dpp_2015 = dpp_2015 + cek_Money(rs(39))
            dpp_2016 = dpp_2016 + cek_Money(rs(41))
            dpp_2017 = dpp_2017 + cek_Money(rs(43))
            dpp_2018 = dpp_2018 + cek_Money(rs(45))
            dpp_2019 = dpp_2019 + cek_Money(rs(47))
            dpp_2020 = dpp_2020 + cek_Money(rs(49))
            total_dpp_all = total_dpp_all + cek_Money(rs(50))
            rs.MoveNext
        Loop
        
        '-- ada di belakang
        rs.AddNew
        rs.Fields(9).Value = "SUM:"
        rs.Fields(10).Value = PU_2008
        rs.Fields(11).Value = PU_2009
        rs.Fields(12).Value = PU_2010
        rs.Fields(13).Value = PU_2011
        rs.Fields(14).Value = PU_2012
        rs.Fields(15).Value = PU_2013
        rs.Fields(16).Value = PU_2014
        rs.Fields(17).Value = PU_2015
        rs.Fields(18).Value = PU_2016
        rs.Fields(19).Value = PU_2017
        rs.Fields(20).Value = PU_2018
        rs.Fields(21).Value = PU_2019
        rs.Fields(22).Value = PU_2020
        rs.Fields(23).Value = Jumlah
        
        rs.Fields(25).Value = dpp_2008
        rs.Fields(27).Value = dpp_2009
        rs.Fields(29).Value = dpp_2010
        rs.Fields(31).Value = dpp_2011
        rs.Fields(33).Value = dpp_2012
        rs.Fields(35).Value = dpp_2013
        rs.Fields(37).Value = dpp_2014
        rs.Fields(39).Value = dpp_2015
        rs.Fields(41).Value = dpp_2016
        rs.Fields(43).Value = dpp_2017
        rs.Fields(45).Value = dpp_2018
        rs.Fields(47).Value = dpp_2019
        rs.Fields(49).Value = dpp_2020
        rs.Fields(50).Value = total_dpp_all
        rs.Update
        
        Call format_Grid
        Call info(1, "Jumlah data : " & jRec, Me.StatusBar1)
    End If
    Me.Enable_Form
End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
  Call dbMySQL_open
    
  'load combo
  Call tbAll2016_loadDivisi(Me.cb_divisi)
  Me.cb_proyek.text = ""
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
    Me.cb_divisi.Enabled = True
  'Else
  '  Call pesan2("Level tidak valid", , vbYellow)
 '   Me.cb_divisi.Enabled = False
 ' End If
  
End Sub


Private Sub Form_Resize()
    If Me.Width - 375 > 0 Then Me.Frame3.Width = Me.Width - 375
    If Me.Height - 2865 > 0 Then Me.Frame3.Height = Me.Height - 2865
    
    If Me.Width - 615 > 0 Then Me.DataGrid1.Width = Me.Width - 615
    If Me.Height - 3705 > 0 Then Me.DataGrid1.Height = Me.Height - 3705

    If Me.Height - 3360 > 0 Then Me.txt_cari.Top = Me.Height - 3360
    Me.Label6.Top = Me.txt_cari.Top
    Me.cmd_export.Top = Me.txt_cari.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call dbMySQL_close
End Sub

Private Sub txt_cari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmd_Load_Click
    End If
End Sub
