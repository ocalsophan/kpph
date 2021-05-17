VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_repKaryawan 
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
      Caption         =   "Report Karyawan"
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
Attribute VB_Name = "frm_repKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset

Sub load_grid(Optional sql As String = "")
    Dim jRec As Long
    Dim c As Long
    Dim rsLoad As ADODB.Recordset
    Dim a As Integer
    Dim npwP As String
    
    
    '-- nmKOlom
    '0: kodedivisi, nama_divisi, ket
    '----------
    On Error GoTo er1
    Me.disable_Form
    If dbMySQL_open = False Then
        Call pesan2("Koneksi ke DBOnline gagal", 1)
        Me.Enable_Form
        Exit Sub
    End If
    
    If Trim(sql) = "" Then
        sql = "select NPWP, Nama, Tahun_Pajak, " & _
                "Masa_Pajak, kode_divisi, kd_proyek, " & _
                "Jumlah_Bruto, Jumlah_PPh, '' as status " & _
                "From pph21bulanan " & _
                "where concat(NPWP,ucase(Nama)) in " & _
                "(select concat(npwp,ucase(nama)) from mkaryawan ) " & _
                "order by npwp, tahun_pajak desc, masa_pajak limit 5000"
        'sql = InputBox("", "", sql)
    End If
    
    If OpenRecordSet(cnn, rsLoad, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        MsgBox "error run " & sql, vbCritical
        Exit Sub
    End If
    
    '-- referensi
    '0: NPWP, Nama, Tahun_Pajak, " & _
    '3: Masa_Pajak, kode_divisi, kd_proyek, " & _
    '6: Jumlah_Bruto, Jumlah_PPh, '' as status " & _
    '--
    
    jRec = RecordCount(rsLoad)
    If jRec <= 0 Then
        sql = InputBox("sql", "", sql)
        Call pesan2("tidak ada data")
        Me.Enable_Form
        Exit Sub
    Else
        
        'copy rsLOAD to RS
        If createRS_duplicate(rsLoad, rs) = True Then
            rsLoad.MoveFirst
            c = 1
            Do While rsLoad.EOF = False
                Call info_progress(Me.StatusBar1, 1, c, jRec, "load karyawan")
                rs.AddNew
                For a = 0 To rsLoad.Fields.Count - 1
                    If a = 8 Then
                        npwP = cek_null(rsLoad(0))
                        If checkNPWP(npwP) = True Then
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
            
            Set Me.DGrid1.DataSource = rs
            For c = 0 To rs.Fields.Count - 1
                'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
       
                If c = 6 Or c = 7 Then
                    Me.DGrid1.Columns(c).Alignment = dbgRight
                    Me.DGrid1.Columns(c).NumberFormat = "###,###"
                    Me.DGrid1.Columns(c).Width = 1400
                End If
            Next
        End If
        
        Me.Frame3.Caption = "Jumlah Data: " & jRec
    End If
    
    Me.Frame3.Caption = "Load top 5000"
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




Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
    
  Call load_grid
  Me.txt_cari.Text = ""
      
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set Me.DGrid1.DataSource = Nothing
    Set rs = Nothing
    Call dbMySQL_close
End Sub

Private Sub txt_cari_KeyPress(KeyAscii As Integer)
    Dim sql As String
    
    If KeyAscii = 13 Then
        sql = "select NPWP, Nama, Tahun_Pajak, " & _
                "Masa_Pajak, kode_divisi, kd_proyek, " & _
                "Jumlah_Bruto, Jumlah_PPh, '' as status " & _
                "From pph21bulanan " & _
                "where concat(NPWP,ucase(Nama)) in " & _
                "(select concat(npwp,ucase(nama)) from mnpwp where skaryawan = 1) " & _
                "and (NPWP like '%" & cleanStr(Me.txt_cari.Text) & "%' or " & _
                "Nama like '%" & cleanStr(Me.txt_cari.Text) & "%' or " & _
                "kd_proyek like '%" & cleanStr(Me.txt_cari.Text) & "%')" & _
                "order by npwp, tahun_pajak desc, masa_pajak"
        sql = "select NPWP, Nama, Tahun_Pajak, " & _
                "Masa_Pajak, kode_divisi, kd_proyek, " & _
                "Jumlah_Bruto, Jumlah_PPh, '' as status " & _
                "From pph21bulanan " & _
                "where concat(NPWP,ucase(Nama)) in " & _
                "(select concat(npwp,ucase(nama)) from mnpwp) " & _
                "and (NPWP like '%" & cleanStr(Me.txt_cari.Text) & "%' or " & _
                "Nama like '%" & cleanStr(Me.txt_cari.Text) & "%' or " & _
                "kd_proyek like '%" & cleanStr(Me.txt_cari.Text) & "%')" & _
                "order by npwp, tahun_pajak desc, masa_pajak"
                
        Call load_grid(sql)
    End If
End Sub
