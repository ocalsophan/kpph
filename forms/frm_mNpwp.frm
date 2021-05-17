VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_mNpwp 
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
   Begin VB.Frame Frame3 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   12015
      Begin VB.CommandButton cmd_setStatus 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Set Karyawan"
         Height          =   375
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton cmd_hapus1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Hapus Data(s)"
         Height          =   375
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5760
         Width           =   1695
      End
      Begin VB.CommandButton cmd_export 
         Caption         =   "&Export XLS"
         Height          =   375
         Left            =   10920
         TabIndex        =   6
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox txt_cari 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Text            =   "Text1"
         ToolTipText     =   "input dan ENTER"
         Top             =   5760
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid DGrid1 
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
      Caption         =   "Master NPWP WP"
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
Attribute VB_Name = "frm_mNpwp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset

Sub load_grid(Optional sql As String = "")
    Dim jRec As Long
    
    '-- nmKOlom
    '0: kodedivisi, nama_divisi, ket
    '----------
    On Error GoTo er1
    Me.disable_Form
    Call dbMySQL_open
    
    If Trim(sql) = "" Then
        sql = "select npwp, nama, alamat, skaryawan from mnpwp order by npwp limit 500"
    End If
    
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        MsgBox "error run " & sql, vbCritical
        Me.Enable_Form
        Exit Sub
    End If
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then
        Call pesan2("tidak ada data")
        Me.Enable_Form
        Exit Sub
    End If
    
    Set Me.DGrid1.DataSource = rs
    Me.Frame3.Caption = "Load top 500"
    Me.Enable_Form
    Exit Sub
er1:
    MsgBox Err.Description, vbCritical
    Me.Enable_Form
End Sub





Sub disable_Form()
    Me.Frame3.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame3.Enabled = True
End Sub





Private Sub cmd_export_Click()
    Dim jRec As Long
    
    Me.disable_Form
    jRec = RecordCount(rs)
    If jRec > 0 Then
        Call create_xls2(rs, "Master NPWP WP", "", "")
    End If
    Me.Enable_Form
End Sub

Private Sub cmd_hapus1_Click()
    Dim j As Integer, rec_no As Long
    Dim npwp_wp As String
    Dim p
    Dim isAdaYangDihapus As Boolean
    
    On Error GoTo er1
    isAdaYangDihapus = False
    
    For j = 0 To Me.DGrid1.SelBookmarks.Count - 1
        rec_no = Me.DGrid1.SelBookmarks.Item(j)
        rs.AbsolutePosition = rec_no
        
        npwp_wp = cek_null(rs(0))
        p = MsgBox("Yakin menghapus 1 record data untuk " & vbCr & "NPWP: " & npwp_wp & vbCr & _
                    "?", vbYesNo)
        If p = vbYes Then
            
            If tbMNpwp_Delete(npwp_wp) = True Then
                isAdaYangDihapus = True
            Else
                Call pesan2("Hapus gagal", 100, vbYellow)
            End If
        End If
    Next
    
    If isAdaYangDihapus = True Then Call txt_cari_KeyPress(13)
    Exit Sub
er1:
    MsgBox Err.Description, vbCritical
End Sub


Private Sub cmd_setStatus_Click()
    Dim j As Integer, rec_no As Long
    Dim npwp_wp As String
    Dim p
    Dim isAdaYangDiUbah As Boolean
    
    On Error GoTo er1
    isAdaYangDiUbah = False
    
    For j = 0 To Me.DGrid1.SelBookmarks.Count - 1
        rec_no = Me.DGrid1.SelBookmarks.Item(j)
        rs.AbsolutePosition = rec_no
        
        npwp_wp = cek_null(rs(0))
        If tbMNpwp_setStatus(npwp_wp) = True Then
            isAdaYangDiUbah = True
        Else
            Exit Sub
        End If
    Next
    
    If isAdaYangDiUbah = True Then Call txt_cari_KeyPress(13)
    Exit Sub
er1:
    MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
  Call load_grid
  Me.txt_cari.Text = ""
      
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set Me.DGrid1.DataSource = Nothing
    Call dbMySQL_close
End Sub

Private Sub txt_cari_KeyPress(KeyAscii As Integer)
    Dim sql As String
    
    If KeyAscii = 13 Then
        sql = "select npwp, nama, alamat, skaryawan from mnpwp where npwp like '%" & cleanStr(Me.txt_cari.Text) & _
                "%' or nama like '%" & cleanStr(Me.txt_cari.Text) & "%' order by npwp limit 500"
        Call load_grid(sql)
    End If
End Sub
