VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_EkFP_rekap 
   ClientHeight    =   7245
   ClientLeft      =   240
   ClientTop       =   750
   ClientWidth     =   12360
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
   ScaleWidth      =   12360
   Begin VB.Frame Frame1 
      Caption         =   " Filter "
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   12015
      Begin VB.CommandButton cmd_load 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Load"
         Height          =   375
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cb_tahun 
         Height          =   330
         Left            =   840
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   382
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun "
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   442
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   12015
      Begin VB.CommandButton cmd_export 
         Caption         =   "&Export XLS"
         Height          =   375
         Left            =   10920
         TabIndex        =   4
         Top             =   4920
         Width           =   975
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
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6990
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10610
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10610
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
      Caption         =   "Ekualisasi : Data FP / Faktur Pajak - Rekap"
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
Attribute VB_Name = "frm_EkFP_rekap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim nama_data As String


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
    kondisi = ""
    cari = ""
        
    sql = "select tahun, kode_proyek_lama,  pk_pm, sum(dpp), sum(ppn) " & _
            "From all2016_fp "
    
    If Not (Trim(Me.cb_tahun) = "" Or Trim(Me.cb_tahun) = "ALL") Then
        kondisi = "tahun = '" & Trim(Me.cb_tahun) & "' "
    End If

    
    '-- gabungkan kondisi
    If Trim(kondisi) <> "" Then
        sql = sql & " where (" & kondisi & ") "
    End If
    
    sql = sql & "group by tahun, kode_proyek_lama, pk_pm "
        
    generate_sql = sql
    'Me.Frame3.Caption = " " & kondisi & " / " & Trim(Me.txt_cari.text)
End Function

Sub format_Grid()
    
    Dim jRec As Long
    Dim c As Integer
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
            
            'kecil
            If c = 0 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).Width = 400
            End If
            
            'If c = 12 Or c = 20 Then
            '    Me.DataGrid1.Columns(c).Alignment = dbgCenter
            '    Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
            '    Me.DataGrid1.Columns(c).Width = 900
            'End If
    
            If c = 2 Or c = 3 Then
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


Private Sub cmd_load_Click()
    Call LoadGrid
End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
  nama_data = "Trial Balance"
  Call dbMySQL_open
    
  'load combo
  sql = "select distinct tahun from all2016_fp"
  Call Load_combo(Me.cb_tahun, sql, cnn, True, , 0)
  
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
 
End Sub


Private Sub Form_Resize()
    If Me.Width - 405 > 0 Then Me.Frame3.Width = Me.Width - 405
    If Me.Height - 2595 > 0 Then Me.Frame3.Height = Me.Height - 2595

    If Me.Width - 645 > 0 Then Me.DataGrid1.Width = Me.Width - 645
    If Me.Height - 3435 > 0 Then Me.DataGrid1.Height = Me.Height - 3435

    If Me.Height - 3090 > 0 Then Me.cmd_export.Top = Me.Height - 3090
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call dbMySQL_close
End Sub


