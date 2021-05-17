VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_csv_BNI_rep 
   ClientHeight    =   7380
   ClientLeft      =   300
   ClientTop       =   810
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
   ScaleHeight     =   7380
   ScaleWidth      =   12300
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4455
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   12015
      Begin VB.CommandButton cmd_xls 
         Caption         =   "Export XLS"
         Height          =   375
         Left            =   10680
         TabIndex        =   13
         Top             =   3960
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3615
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   6376
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
   Begin VB.CommandButton cmd_proses 
      Caption         =   "2. Load"
      Height          =   375
      Left            =   10560
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
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
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   12015
      Begin VB.ListBox List1 
         Height          =   900
         Left            =   6600
         TabIndex        =   12
         Top             =   240
         Width           =   5295
      End
      Begin VB.TextBox txt_masa 
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txt_Tahun 
         Height          =   375
         Left            =   4560
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cb_divisi 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Text            =   "x"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Line Line2 
         X1              =   6360
         X2              =   6360
         Y1              =   240
         Y2              =   1080
      End
      Begin VB.Line Line1 
         X1              =   3960
         X2              =   3960
         Y1              =   240
         Y2              =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Masa"
         Height          =   210
         Left            =   4080
         TabIndex        =   8
         Top             =   780
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   4080
         TabIndex        =   7
         Top             =   420
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Divisi"
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   420
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   7125
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
      Caption         =   "Tax Inquiry Report"
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
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frm_csv_BNI_rep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsGrid As ADODB.Recordset


Sub disable_Form()
    Me.Frame1.Enabled = False
    Me.Frame2.Enabled = False
    Me.cmd_proses.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame1.Enabled = True
    Me.Frame2.Enabled = True
    Me.cmd_proses.Enabled = True
End Sub


Private Sub cmd_proses_Click()
    Call LoadGrid
End Sub

Private Sub LoadGrid()
    Dim sql As String, jRec As Long
    
    Me.disable_Form
    sql = generate_sql
    DoEvents
    
    If Trim(sql) <> "" Then
        Call dbMySQL_open
        If OpenRecordSet(cnn, rsGrid, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
            sql = InputBox("", "", sql)
            MsgBox "error open sql"
            Me.Enable_Form
            Exit Sub
        End If
        
        Set Me.DataGrid1.DataSource = rsGrid
        jRec = RecordCount(rsGrid)
    End If
    Call format_Grid
    Me.Enable_Form
End Sub


Function generate_sql() As String
    Dim sql As String, kondisi As String
    Dim cari As String
    
    'kondisi
    kondisi = ""
        
    sql = "Select tax_inquiry.trans_reference_no, tax_inquiry.created_date, tax_inquiry.transaction_date, " & _
        "tax_inquiry.posting_date, tax_inquiry.billing_id, tax_inquiry.NTPN, " & _
        "tax_inquiry.NPWP_number, tax_inquiry.Customer_Reference_No, tax_inquiry.Remark, " & _
        "tax_inquiry.Currency1, tax_inquiry.Amount, tax_inquiry.Status, " & _
        "tax_inquiry.Reason, log_export_mandiri.kd_divisi, log_export_mandiri.tgl, " & _
        "log_export_mandiri.nmFile,log_export_mandiri.k10_Tahun_paj, log_export_mandiri.k8_Masa_paja " & _
        "From " & _
        "tax_inquiry Left Join " & _
        "log_export_mandiri On log_export_mandiri.k14_Customer_ = tax_inquiry.Customer_Reference_No"
    
    If Not (Trim(Me.cb_divisi.text) = "" Or UCase(Trim(Me.cb_divisi.text)) = "ALL") Then
        kondisi = "kd_divisi = '" & Trim(get_kode_combo(Me.cb_divisi, "-")) & "' "
    End If
    
    If Trim(Me.txt_Tahun.text) <> "" Then
        If Trim(kondisi) = "" Then
            kondisi = "k10_Tahun_paj = '" & Trim(Me.txt_Tahun) & "' "
        Else
            kondisi = kondisi & " and k10_Tahun_paj = '" & Trim(Me.txt_Tahun) & "' "
        End If
    End If
    
    If Trim(Me.txt_masa.text) <> "" Then
        If Trim(kondisi) = "" Then
            kondisi = "k8_Masa_paja = '" & Trim(Me.txt_masa) & "' "
        Else
            kondisi = kondisi & " and k8_Masa_paja = '" & Trim(Me.txt_masa) & "' "
        End If
    End If
    
    '-- ini sql cari
    'If Trim(Me.txt_cari.text) <> "" Then
    '    cari = "kode_proyek_baru like '%" & Trim(Me.txt_cari.text) & "%' or " & _
    '            "no_fp like '%" & Trim(Me.txt_cari.text) & "%' or " & _
    '            "no_bp like '%" & Trim(Me.txt_cari.text) & "%' or " & _
    '            "keterangan like '%" & Trim(Me.txt_cari.text) & "%' "
    'End If
    
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
        
    sql = sql & " order by created_date "
        
    generate_sql = sql
End Function

Sub format_Grid()
    
    Dim jRec As Long
    Dim c As Integer
    
    jRec = RecordCount(rsGrid)
    If jRec <= 0 Then Exit Sub
    
        For c = 0 To rsGrid.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
            
            'kecil
            'If c = 0 Or c = 1 Then
            '    Me.DataGrid1.Columns(c).Alignment = dbgCenter
            '    Me.DataGrid1.Columns(c).Width = 400
            'End If
            
            'If c = 12 Or c = 20 Then
            '    Me.DataGrid1.Columns(c).Alignment = dbgCenter
            '    Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
            '    Me.DataGrid1.Columns(c).Width = 900
            'End If
    
            If c = 10 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
End Sub

Private Sub cmd_xls_Click()
    Call create_xls2(rsGrid, "", "", "")
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
  Call load_Divisi(Me.cb_divisi, False, 1, True)
  
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
  
  Me.txt_Tahun.text = Year(Now)
  Me.txt_masa.text = Month(Now)
  
  
  Me.Width = 12540
  Me.Height = 8175
  Me.cmd_xls.Top = 3960
  Me.cmd_xls.Left = 10680
End Sub


Private Sub Form_Resize()
    Me.Shape1.Width = Me.Width
    Me.lb_caption.Width = Me.Width
    
    If Me.Width - 645 > 0 Then Me.Frame2.Width = Me.Width - 645
    If Me.Frame2.Width - 240 > 0 Then Me.DataGrid1.Width = Me.Frame2.Width - 240
    
    'height
    If Me.Height - 3720 > 0 Then Me.Frame2.Height = Me.Height - 3720
    If Me.Frame2.Height - 840 > 0 Then Me.DataGrid1.Height = Me.Frame2.Height - 840
    
    If Frame2.Height - 495 > 0 Then Me.cmd_xls.Top = Frame2.Height - 495
    'If Me.Width - 1740 > 0 Then Me.cmd_xls.Left = Me.Width - 1740
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call dbMySQL_close
End Sub

