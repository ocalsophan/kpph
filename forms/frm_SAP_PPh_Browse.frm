VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_SAP_PPh_Browse 
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
      TabIndex        =   10
      Top             =   2040
      Width           =   12015
      Begin VB.CommandButton cmd_hapus1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Hapus Data(s)"
         Height          =   375
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CommandButton cmd_export 
         Caption         =   "&Export XLS"
         Height          =   375
         Left            =   10920
         TabIndex        =   14
         Top             =   4320
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3975
         Left            =   120
         TabIndex        =   12
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
      Begin VB.Label lb_total 
         AutoSize        =   -1  'True
         Caption         =   "--"
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   4395
         Width           =   120
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " 2. Masa / Posting Key"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   600
      Width           =   6375
      Begin VB.OptionButton opt_rekap 
         Caption         =   "Rekap"
         Height          =   255
         Left            =   5280
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton opt_detil 
         Caption         =   "Detil"
         Height          =   255
         Left            =   4440
         TabIndex        =   18
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.ComboBox cb_posting_key 
         Height          =   330
         Left            =   1440
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cb_year_month 
         Height          =   330
         Left            =   1440
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Load 
         Caption         =   "3. &Load"
         Height          =   375
         Left            =   5040
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Posting Key"
         Height          =   210
         Left            =   360
         TabIndex        =   9
         Top             =   915
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Year Month"
         Height          =   210
         Left            =   360
         TabIndex        =   8
         Top             =   435
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 1. Account / Profit Center"
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
      Width           =   5535
      Begin VB.ComboBox cb_jenisPajak 
         Height          =   330
         Left            =   1200
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   720
         Width           =   4215
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
         Caption         =   "Profit Center"
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Account"
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   780
         Width           =   615
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
      Caption         =   "Browse Data PPh SAP"
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
Attribute VB_Name = "frm_SAP_PPh_Browse"
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
    If t = "1" Or t = "2" Or t = "3" Or t = "4" Or t = "5" Or t = "6" Or t = "7" Or t = "8" Or t = "9" Then
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



Sub disable_Form()
    Me.Frame1.Enabled = False
    Me.Frame2.Enabled = False
    Me.Frame3.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame1.Enabled = True
    Me.Frame2.Enabled = True
    Me.Frame3.Enabled = True
End Sub


Function generate_sql() As String
    Dim sql As String, divisi As String
    
    If UCase(Trim(Me.cb_divisi)) = "ALL" Then
        divisi = ""
    Else
        divisi = Left(Me.cb_divisi.text, 2)
    End If
    
    If Me.opt_detil.Value = True Then
        Call pesan2("Detil", 1)
        sql = "call P_pph_sap_rep('" & get_kode_combo(Me.cb_jenisPajak, "-") & _
            "','" & divisi & "','" & Me.cb_year_month & "','" & _
            Me.cb_posting_key & "')"
    ElseIf Me.opt_rekap.Value = True Then
        Call pesan2("Rekap", 1)
        sql = "call P_pph_sap_rep2('" & divisi & "','" & Me.cb_year_month & "','" & _
            Me.cb_posting_key & "')"
    End If
    generate_sql = sql
End Function

Sub format_Grid()
    
    Dim jenisPPh As String
    Dim jRec As Long
    Dim c As Integer
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
    If Me.opt_detil.Value = True Then
    
        For c = 0 To rs.Fields.Count - 1
            If c = 0 Then
                Me.DataGrid1.Columns(c).Width = 1000
            ElseIf c = 1 Or c = 2 Then
                Me.DataGrid1.Columns(c).Width = 600
            ElseIf c = 5 Or c = 6 Then
                 Me.DataGrid1.Columns(c).Width = 700
            ElseIf c = 7 Or c = 8 Then
                Me.DataGrid1.Columns(c).Width = 1000
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
    
            'If c = 12 Or c = 20 Then
            '    Me.DataGrid1.Columns(c).Alignment = dbgCenter
            '    Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
            '    Me.DataGrid1.Columns(c).Width = 900
            'End If
    
            ElseIf c = 9 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
    ElseIf Me.opt_rekap.Value = True Then
        For c = 0 To rs.Fields.Count - 1
            If c = 0 Then
                Me.DataGrid1.Columns(c).Width = 700
            ElseIf c = 1 Or c = 3 Then
                Me.DataGrid1.Columns(c).Width = 1000
            ElseIf c = 2 Then
                Me.DataGrid1.Columns(c).Width = 1700
            ElseIf c = 3 Then
                Me.DataGrid1.Columns(c).Width = 700
            ElseIf c = 4 Or c = 8 Then
                Me.DataGrid1.Columns(c).Width = 600
            ElseIf c = 5 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
    End If
End Sub


Private Sub cmd_export_Click()
    Dim jRec As Long
    
    Me.disable_Form
    jRec = RecordCount(rs)
    If jRec > 0 Then
        If Me.opt_detil.Value = True Then
            Call create_xls2(rs, "Detil " & Me.cb_jenisPajak, "09", "")
        ElseIf Me.opt_rekap.Value = True Then
            Call create_xls2(rs, "Rekap " & Me.cb_divisi, "05", "")
        End If
    End If
    Me.Enable_Form
End Sub


Private Sub cmd_hapus1_Click()
    Dim j As Integer, rec_no As Long
    Dim docnumber As String, amount_in_lc  As Currency, id1 As String
    Dim p
    Dim indexAkhir As Integer
    Dim isAdaYangDihapus As Boolean
    
    On Error GoTo er1
    
    If Me.opt_detil.Value = True Then
        Call pesan2("tidak dapat menghapus data rekap")
        Exit Sub
    End If
    
    isAdaYangDihapus = False
    indexAkhir = rs.Fields.Count - 1
    For j = 0 To Me.DataGrid1.SelBookmarks.Count - 1
        rec_no = Me.DataGrid1.SelBookmarks.Item(j)
        rs.AbsolutePosition = rec_no
        
        docnumber = cek_null(rs(0))
        amount_in_lc = cek_Money(cek_null(rs(9)))
        id1 = cek_null(rs(indexAkhir))
        p = MsgBox("Yakin menghapus 1 record data untuk " & vbCr & docnumber & vbCr & _
                    "nilai : " & Format(amount_in_lc, "###,###") & vbCr & "?", vbYesNo)
        If p = vbYes Then
            isAdaYangDihapus = True
            Call tbPph_sap_delete(id1)
        End If
    Next
    
    If isAdaYangDihapus = True Then Call cmd_Load_Click
    
    Exit Sub
er1:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmd_Load_Click()
    Dim sql As String, jRec As Long
    Dim tot_amount_in_lc As Currency
    
    On Error GoTo er1
    
    Me.disable_Form
    sql = generate_sql
    DoEvents
    'MsgBox sql
    If Trim(sql) <> "" Then
        Call dbMySQL_open
        If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
            MsgBox "error open sql"
            Me.Enable_Form
            Exit Sub
        End If
        
        If Me.opt_detil.Value = True Then
            'get total
            tot_amount_in_lc = 0
            rs.MoveFirst
            Do While rs.EOF = False
                tot_amount_in_lc = tot_amount_in_lc + cek_Money(rs(9))
                rs.MoveNext
            Loop
            Me.lb_total.Caption = "Jumlah amount_in_LC: " & Format(tot_amount_in_lc, "###,###")
        Else
            Me.lb_total.Caption = "-"
        End If
        Set Me.DataGrid1.DataSource = rs
        jRec = RecordCount(rs)
        Call format_Grid
        Call info(1, "Jumlah data : " & jRec, Me.StatusBar1)
        
    End If
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
  Call load_Divisi(Me.cb_divisi, , 1)
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
    
  Call tbPph_sap_load_Account(Me.cb_jenisPajak)
  Call tbPph_sap_load_year_month1(Me.cb_year_month)
  Call tbPph_sap_load_posting_key(Me.cb_posting_key)
  
  Me.Width = 12420
  Me.Height = 7710
End Sub


Private Sub Form_Resize()
    Me.Shape1.Width = Me.Width
    Me.lb_caption.Width = Me.Width
    
    If Me.Width - 405 > 0 Then Me.Frame3.Width = Me.Width - 405
    If Me.Frame3.Width - 240 > 0 Then Me.DataGrid1.Width = Me.Frame3.Width - 240
    
    'Height
    If Me.Height - 2895 > 0 Then Me.Frame3.Height = Me.Height - 2895
    If Me.Frame3.Height - 840 > 0 Then Me.DataGrid1.Height = Me.Frame3.Height - 840
    
    If Me.Frame3.Height - 495 > 0 Then Me.cmd_hapus1.Top = Me.Frame3.Height - 495
    Me.cmd_export.Top = Me.cmd_hapus1.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call dbMySQL_close
End Sub

Sub cek_rekap()
    If Me.opt_detil.Value = True Then
        Me.cb_jenisPajak.Enabled = True
    ElseIf Me.opt_rekap.Value = True Then
        Me.cb_jenisPajak.Enabled = False
    Else
        Me.cb_jenisPajak.Enabled = False
    End If
End Sub

Private Sub opt_detil_Click()
    Call cek_rekap
End Sub

Private Sub opt_rekap_Click()
    Call cek_rekap
End Sub
