VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_SPTBadan 
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
      TabIndex        =   7
      Top             =   480
      Width           =   12015
      Begin VB.ComboBox cb_masa 
         Height          =   330
         Left            =   2760
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cb_tahun 
         Height          =   330
         Left            =   840
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd_load 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Load"
         Height          =   375
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Masa"
         Height          =   210
         Left            =   2280
         TabIndex        =   10
         Top             =   420
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun "
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   420
         Visible         =   0   'False
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
      Caption         =   "SPT Badan"
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
         Caption         =   "Hapus Data"
      End
   End
End
Attribute VB_Name = "frm_SPTBadan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim nama_data As String
Dim isDataBerubah As Boolean


Sub disable_Form()
    Me.Frame3.Enabled = False
    Me.Frame1.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame3.Enabled = True
    Me.Frame1.Enabled = True
End Sub


Sub spt_badan_SimpanData()
    Dim jRec As Long, c As Long
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    rs.MoveFirst
    c = 1
    Do While rs.EOF = False
        Call info_progress(Me.StatusBar1, 1, c, jRec, "Update Data")
        
        If Trim(cek_null(rs(0))) <> "" Then
            Call tbVariabel_set(cek_null(rs(0)), cek_null(rs(1)))
        End If
        
        c = c + 1
        rs.MoveNext
    Loop
End Sub


Sub spt_badan_default()
    '-- set variabel default
    
    Dim var1
    Dim a As Integer, t As String
    
    var1 = Array("sptbadan_0_tahun", "sptbadan_0_npwp", "sptbadan_0_namawp", _
                "sptbadan_0_jenisusaha", "sptbadan_0_klu", "sptbadan_0_telepon", _
                "sptbadan_0_faks", "sptbadan_0_periodebuku1", "sptbadan_0_periodebuku2", _
                "sptbadan_0_negaradomisili", "sptbadan_0_pembukulanlaporan", _
                "sptbadan_0_namakantorakuntan", "sptbadan_0_npwpakantorkuntan", _
                "sptbadan_0_namaakuntan", "sptbadan_0_npwpakuntan", _
                "sptbadan_0_namakantorkonsultan", "sptbadan_0_npwpkantorkonsultan", _
                "sptbadan_0_namakonsultan", "sptbadan_0_npwpkonsultan")
    For a = 0 To UBound(var1)
        Call info_progress(Me.StatusBar1, 1, CLng(a), CLng(UBound(var1)), "Set default variabel")
        t = tbVariabel_get(CStr(var1(a)))
        If Trim(t) = "" Then
            Call tbVariabel_set(CStr(var1(a)), "-")
        End If
    Next
    
    '-- load di rsgrid
    Call create_rs2(rs, "key1; value1")
    
    For a = 0 To UBound(var1)
        Call info_progress(Me.StatusBar1, 1, CLng(a), CLng(UBound(var1)), "Set default variabel")
        rs.AddNew
        t = tbVariabel_get(CStr(var1(a)))
        rs.Fields(0) = CStr(var1(a))
        rs.Fields(1) = t
        rs.Update
    Next
    Set Me.DataGrid1.DataSource = rs
    Me.DataGrid1.Columns(0).Locked = True
    
End Sub


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




Private Sub cmd_export_Click()
    Dim p

    On Error GoTo er1
    If isDataBerubah = True Then
        p = MsgBox("Simpan perubahan ? ", vbYesNo)
        If p = vbYes Then
            '--- simpan perubahan
            Call spt_badan_SimpanData
            isDataBerubah = False
        End If
    End If
    
    
    Call proses_xls
    MsgBox "done"
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
End Sub

Sub proses_xls_sheet1(ByRef fl As Object)
    Dim fLs As Object
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim txt1 As String, t1 As String
    
    'open sheet 1, isi data looping
    Set fLs = fl.Sheets(1)
    
    'tahun
    baris = 4
    kolom = 33
    txt1 = tbVariabel_get("sptbadan_0_tahun")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        kolom = kolom + 2
    Next
    
    
    'npwp
    baris = 11
    kolom = 10
    txt1 = tbVariabel_get("sptbadan_0_npwp")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        If a = 2 Or a = 5 Or a = 8 Or a = 9 Or a = 12 Then
            kolom = kolom + 2
        Else
            kolom = kolom + 1
        End If
    Next
    
    'sptbadan_0_namawp
    baris = 13
    kolom = 10
    txt1 = tbVariabel_get("sptbadan_0_namawp")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        kolom = kolom + 1
    Next
    
    'sptbadan_0_jenisusaha
    baris = 15
    kolom = 10
    txt1 = tbVariabel_get("sptbadan_0_jenisusaha")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        kolom = kolom + 1
    Next
    
    'sptbadan_0_klu
    baris = 15
    kolom = 34
    txt1 = tbVariabel_get("sptbadan_0_klu")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        kolom = kolom + 1
    Next
    
    'sptbadan_0_telepon
    baris = 17
    kolom = 10
    txt1 = tbVariabel_get("sptbadan_0_telepon")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        If a = 4 Then
            kolom = kolom + 2
        Else
            kolom = kolom + 1
        End If
    Next
    
    'sptbadan_0_faks
    baris = 17
    kolom = 27
    txt1 = tbVariabel_get("sptbadan_0_faks")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        If a = 4 Then
            kolom = kolom + 2
        Else
            kolom = kolom + 1
        End If
    Next
    
    
    'sptbadan_0_periodebuku1
    baris = 19
    kolom = 10
    txt1 = tbVariabel_get("sptbadan_0_periodebuku1")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
    'sptbadan_0_periodebuku2
    baris = 19
    kolom = 17
    txt1 = tbVariabel_get("sptbadan_0_periodebuku2")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
    'sptbadan_0_negaradomisili
    baris = 21
    kolom = 17
    txt1 = tbVariabel_get("sptbadan_0_negaradomisili")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
    'sptbadan_0_pembukulanlaporan
    baris = 25
    txt1 = tbVariabel_get("sptbadan_0_pembukulanlaporan")
    If UCase(Trim(txt1)) = "DIAUDIT" Then
        fLs.Cells(baris, 13) = "X"
    ElseIf UCase(Trim(txt1)) = "OPINI AKUNTAN" Then
        fLs.Cells(baris, 18) = "X"
    Else
        fLs.Cells(baris, 24) = "X"
    End If
    
    'sptbadan_0_namakantorakuntan
    baris = 27
    kolom = 13
    txt1 = tbVariabel_get("sptbadan_0_namakantorakuntan")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
    'sptbadan_0_npwpakantorkuntan
    baris = 29
    kolom = 13
    txt1 = tbVariabel_get("sptbadan_0_npwpakantorkuntan")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        If a = 2 Or a = 5 Or a = 8 Or a = 9 Or a = 12 Then
            kolom = kolom + 2
        Else
            kolom = kolom + 1
        End If
    Next
    
    
    '---
    'sptbadan_0_namaakuntan
    baris = 31
    kolom = 13
    txt1 = tbVariabel_get("sptbadan_0_namaakuntan")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
    'sptbadan_0_npwpakuntan
    baris = 33
    kolom = 13
    txt1 = tbVariabel_get("sptbadan_0_npwpakuntan")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        If a = 2 Or a = 5 Or a = 8 Or a = 9 Or a = 12 Then
            kolom = kolom + 2
        Else
            kolom = kolom + 1
        End If
    Next
    
    '---
    'sptbadan_0_namakantorkonsultan
    baris = 35
    kolom = 13
    txt1 = tbVariabel_get("sptbadan_0_namakantorkonsultan")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
    'sptbadan_0_npwpkantorkonsultan
    baris = 37
    kolom = 13
    txt1 = tbVariabel_get("sptbadan_0_npwpkantorkonsultan")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        If a = 2 Or a = 5 Or a = 8 Or a = 9 Or a = 12 Then
            kolom = kolom + 2
        Else
            kolom = kolom + 1
        End If
    Next
    
    '---
    'sptbadan_0_namakonsultan
    baris = 39
    kolom = 13
    txt1 = tbVariabel_get("sptbadan_0_namakonsultan")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
    'sptbadan_0_npwpkonsultan
    baris = 41
    kolom = 13
    txt1 = tbVariabel_get("sptbadan_0_npwpkonsultan")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        If a = 2 Or a = 5 Or a = 8 Or a = 9 Or a = 12 Then
            kolom = kolom + 2
        Else
            kolom = kolom + 1
        End If
    Next
    
    
End Sub


Sub proses_xls_sheet3(ByRef fl As Object)
    Dim fLs As Object
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim txt1 As String, t1 As String
    
    'open sheet 1, isi data looping
    Set fLs = fl.Sheets(3)
    
    'tahun
    baris = 5
    kolom = 34
    txt1 = tbVariabel_get("sptbadan_0_tahun")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        kolom = kolom + 2
    Next
    
    
    'npwp
    baris = 10
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_npwp")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        If a = 2 Or a = 5 Or a = 8 Or a = 9 Or a = 12 Then
            kolom = kolom + 2
        Else
            kolom = kolom + 1
        End If
    Next
    
    'sptbadan_0_namawp
    baris = 12
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_namawp")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        kolom = kolom + 1
    Next
        
    
    'sptbadan_0_periodebuku1
    baris = 14
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_periodebuku1")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
    'sptbadan_0_periodebuku2
    baris = 14
    kolom = 17
    txt1 = tbVariabel_get("sptbadan_0_periodebuku2")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
End Sub

Sub proses_xls_sheet4(ByRef fl As Object)
    Dim fLs As Object
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim txt1 As String, t1 As String
    
    'open sheet 1, isi data looping
    Set fLs = fl.Sheets(4)
    
    'tahun
    baris = 4
    kolom = 59
    txt1 = tbVariabel_get("sptbadan_0_tahun")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        kolom = kolom + 2
    Next
    
    
    'npwp
    baris = 9
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_npwp")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        If a = 2 Or a = 5 Or a = 8 Or a = 9 Or a = 12 Then
            kolom = kolom + 2
        Else
            kolom = kolom + 1
        End If
    Next
    
    'sptbadan_0_namawp
    baris = 9
    kolom = 39
    txt1 = tbVariabel_get("sptbadan_0_namawp")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        kolom = kolom + 1
    Next
        
    
    'sptbadan_0_periodebuku1
    baris = 11
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_periodebuku1")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
    'sptbadan_0_periodebuku2
    baris = 11
    kolom = 17
    txt1 = tbVariabel_get("sptbadan_0_periodebuku2")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
End Sub

Sub proses_xls_sheet5(ByRef fl As Object)
    Dim fLs As Object
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim txt1 As String, t1 As String
    
    'open sheet 1, isi data looping
    Set fLs = fl.Sheets(5)
    
    'tahun
    baris = 5
    kolom = 66
    txt1 = tbVariabel_get("sptbadan_0_tahun")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        kolom = kolom + 2
    Next
    
    
    'npwp
    baris = 10
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_npwp")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        If a = 2 Or a = 5 Or a = 8 Or a = 9 Or a = 12 Then
            kolom = kolom + 2
        Else
            kolom = kolom + 1
        End If
    Next
    
    'sptbadan_0_namawp
    baris = 10
    kolom = 39
    txt1 = tbVariabel_get("sptbadan_0_namawp")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        kolom = kolom + 1
    Next
        
    
    'sptbadan_0_periodebuku1
    baris = 12
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_periodebuku1")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
    'sptbadan_0_periodebuku2
    baris = 12
    kolom = 17
    txt1 = tbVariabel_get("sptbadan_0_periodebuku2")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
End Sub

Sub proses_xls_sheet6(ByRef fl As Object)
    Dim fLs As Object
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim txt1 As String, t1 As String
    
    'open sheet 1, isi data looping
    Set fLs = fl.Sheets(6)
    
    'tahun
    baris = 5
    kolom = 34
    txt1 = tbVariabel_get("sptbadan_0_tahun")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        kolom = kolom + 2
    Next
    
    
    'npwp
    baris = 10
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_npwp")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        If a = 2 Or a = 5 Or a = 8 Or a = 9 Or a = 12 Then
            kolom = kolom + 2
        Else
            kolom = kolom + 1
        End If
    Next
    
    'sptbadan_0_namawp
    baris = 12
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_namawp")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        kolom = kolom + 1
    Next
        
    
    'sptbadan_0_periodebuku1
    baris = 14
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_periodebuku1")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
    'sptbadan_0_periodebuku2
    baris = 14
    kolom = 17
    txt1 = tbVariabel_get("sptbadan_0_periodebuku2")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
End Sub

Sub proses_xls_sheet7(ByRef fl As Object)
    Dim fLs As Object
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim txt1 As String, t1 As String
    
    'open sheet 1, isi data looping
    Set fLs = fl.Sheets(7)
    
    'tahun
    baris = 6
    kolom = 34
    txt1 = tbVariabel_get("sptbadan_0_tahun")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        kolom = kolom + 2
    Next
    
    
    'npwp
    baris = 12
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_npwp")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        If a = 2 Or a = 5 Or a = 8 Or a = 9 Or a = 12 Then
            kolom = kolom + 2
        Else
            kolom = kolom + 1
        End If
    Next
    
    'sptbadan_0_namawp
    baris = 14
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_namawp")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        kolom = kolom + 1
    Next
        
    
    'sptbadan_0_periodebuku1
    baris = 16
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_periodebuku1")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
    'sptbadan_0_periodebuku2
    baris = 16
    kolom = 17
    txt1 = tbVariabel_get("sptbadan_0_periodebuku2")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
End Sub

Sub proses_xls_sheet8(ByRef fl As Object)
    Dim fLs As Object
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim txt1 As String, t1 As String
    
    'open sheet 1, isi data looping
    Set fLs = fl.Sheets(8)
    
    'tahun
    baris = 5
    kolom = 34
    txt1 = tbVariabel_get("sptbadan_0_tahun")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        kolom = kolom + 2
    Next
    
    
    'npwp
    baris = 11
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_npwp")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        'fLs.Cells(1, kolom) = t1
        If a = 2 Or a = 5 Or a = 8 Or a = 9 Or a = 12 Then
            kolom = kolom + 2
        Else
            kolom = kolom + 1
        End If
    Next
    
    'sptbadan_0_namawp
    baris = 13
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_namawp")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
        kolom = kolom + 1
    Next
        
    
    'sptbadan_0_periodebuku1
    baris = 15
    kolom = 11
    txt1 = tbVariabel_get("sptbadan_0_periodebuku1")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
    'sptbadan_0_periodebuku2
    baris = 15
    kolom = 17
    txt1 = tbVariabel_get("sptbadan_0_periodebuku2")
    For a = 1 To Len(txt1)
        t1 = Mid(txt1, a, 1)
        fLs.Cells(baris, kolom) = t1
            kolom = kolom + 1
    Next
    
End Sub


Sub proses_xls()
    Dim f As String, nmFile As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim jRec As Long, c As Long
    
    Dim txt1 As String, t1 As String, fileSimpan As String, File1 As String
    
    nmFile = App.Path & "\rep\sptbadan.xls"
    
    fileSimpan = App.Path & "\exp\sptbadan.xls"
                              
                              
    f = nmFile
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Exit Sub
        End If
    Else
        MsgBox "File template tidak ditemukan", vbCritical
        Exit Sub
    End If
    
    Call proses_xls_sheet1(fl)
    Call proses_xls_sheet3(fl)
    Call proses_xls_sheet4(fl)
    Call proses_xls_sheet5(fl)
    Call proses_xls_sheet6(fl)
    Call proses_xls_sheet7(fl)
    Call proses_xls_sheet8(fl)
    
    fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
    
    
    'MsgBox "File tersimpan di " & fileSimpan, vbInformation
    
    'open file
    'open by explorer
    File1 = "explorer.exe " & fileSimpan
    Call Shell(File1, vbNormalFocus)
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
    isDataBerubah = True
End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
  nama_data = "ebupot23"
  Call dbMySQL_open
    
  'load combo
  Me.txt_cari.text = ""
  
  
  
  Me.Height = 8010
  Me.Width = 12420
  
  'get level
  '2 : operator gedung
  '3 : UKP
  
  '---------
  
  'Level1 = tbPengguna_getLevel1(frMenu1.nmLogin)
  'If Level1 = 2 Then
  '  Me.cb_divisi.text = tbPengguna_getDivisi(frMenu1.nmLogin)
  '  Me.cb_divisi.Enabled = False
  'ElseIf Level1 = 3 Then
  'Else
  '  Call pesan2("Level tidak valid", , vbYellow)
  ' Me.cb_divisi.Enabled = False
  'End If
  
  Call spt_badan_default
  isDataBerubah = False
 
End Sub


Private Sub Form_Resize()
    If Me.Width - 405 > 0 Then Me.Frame3.Width = Me.Width - 405
    If Me.Height - 2595 > 0 Then Me.Frame3.Height = Me.Height - 2595

    If Me.Width - 645 > 0 Then Me.DataGrid1.Width = Me.Width - 645
    If Me.Height - 3435 > 0 Then Me.DataGrid1.Height = Me.Height - 3435

    If Me.Height - 3090 > 0 Then Me.txt_cari.Top = Me.Height - 3090
    Me.Label6.Top = Me.txt_cari.Top
    Me.cmd_export.Top = Me.txt_cari.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call dbMySQL_close
End Sub

