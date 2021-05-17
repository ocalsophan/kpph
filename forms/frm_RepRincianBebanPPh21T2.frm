VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_RepRincianBebanPPh21T2 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7155
   ClientLeft      =   225
   ClientTop       =   735
   ClientWidth     =   12900
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
   ScaleHeight     =   7155
   ScaleWidth      =   12900
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   5175
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   12615
      Begin VB.CommandButton cmd_Export 
         Caption         =   "Export"
         Height          =   375
         Left            =   11160
         TabIndex        =   8
         Top             =   4680
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4215
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   7435
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
   Begin VB.Frame Frame1 
      Caption         =   " Filter Data "
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6135
      Begin VB.CommandButton cmd_Load 
         Caption         =   "Load"
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtTahun 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   240
         TabIndex        =   4
         Top             =   442
         Width           =   450
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6900
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11324
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11324
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
      Caption         =   "Rincian Beban -  PPh21 Tahunan"
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
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   12885
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frm_RepRincianBebanPPh21T2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset

Private Sub cmd_export_Click()
    Call create_xls2(rs, "rincian data", "02,03,04,05,06,07,08,09,10", "", "", "", "")
End Sub

Private Sub cmd_Load_Click()
    Dim sql As String
    Dim c As Integer
    
    On Error GoTo er1
    Call dbMySQL_open
    sql = "select kdcenter, kdproyek, sum(nilai_beban) as PPhTerhutang, sum(gaji), sum(Tnj_pph), " & _
            "sum(tunjangan_Lain), sum(JHT_JPN), sum(Bruto), sum(Insentif), " & _
            "Sum(THR), Sum(Lainnya) " & _
            "From pph21tahunan2 " & _
            "where tahun = '" & Trim(Me.txtTahun.Text) & "' " & _
            "group by kdcenter, kdproyek " & _
            "order by kdcenter, kdproyek"
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error run ", "", sql)
        Exit Sub
    End If
    
    Set Me.DataGrid1.DataSource = rs
    
    'format
    For c = 2 To 9
        Me.DataGrid1.Columns(c).Alignment = dbgRight
        Me.DataGrid1.Columns(c).NumberFormat = "###,###"
        '------
    Next
    Exit Sub
er1:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
    Me.txtTahun.Text = Year(Now) - 1
End Sub

