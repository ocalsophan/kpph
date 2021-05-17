VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Browse_CekData 
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
      Height          =   5415
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   12015
      Begin VB.CommandButton cmd_export 
         Caption         =   "&Export XLS"
         Height          =   375
         Left            =   10920
         TabIndex        =   7
         Top             =   4920
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4575
         Left            =   120
         TabIndex        =   6
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
   Begin VB.Frame Frame1 
      Caption         =   " 1. Jenis Cek "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   12015
      Begin VB.CommandButton cmd_Load 
         Caption         =   "Load"
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cb_jenis 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Text            =   "x"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cek"
         Height          =   210
         Left            =   240
         TabIndex        =   4
         Top             =   420
         Width           =   270
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
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
      Caption         =   "Cek Data"
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
Attribute VB_Name = "frm_Browse_CekData"
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
    Dim sql As String
    
    If Left(Me.cb_jenis.Text, 1) = "1" Then
        sql = "select distinct nama, npwp, nik, P_L From pph21tahunan2 order by nama"
    ElseIf Left(Me.cb_jenis.Text, 1) = "2" Then
        sql = "select nama, npwp, nik, npwp_kpp, count(*) " & _
                "From pph21tahunan2 " & _
                "group by nama, npwp, nik, npwp_kpp " & _
                "having count(*) > 12"
    ElseIf Left(Me.cb_jenis.Text, 1) = "3" Then
        sql = "select nama, npwp, nik, count(*) " & _
                "From pph21tahunan2 " & _
                "group by nama, npwp, nik " & _
                "having count(*) < 12"
    ElseIf Left(Me.cb_jenis.Text, 1) = "4" Then
        sql = "select nama, npwp, count(*) " & _
                "from v_nik_double " & _
                "group by nama, npwp " & _
                "having count(*) > 1"
    ElseIf Left(Me.cb_jenis.Text, 1) = "6" Then
        sql = "select nama, npwp, nik, kdcenter, bulan " & _
                "From pph21tahunan2 " & _
                "where nama & npwp & nik in " & _
                "( " & _
                "select nama & npwp & nik " & _
                "From v_jumlah_kurang12 " & _
                ") " & _
                "order by nama, npwp, nik, bulan "
    Else
        sql = ""
    End If
    generate_sql = sql
End Function

Sub load_grid()
    Dim sql As String, t As String
    Dim jRec As Long, c As Long
    Dim nama As String, npwp As String
    Dim NIK As String
    Dim p
    
    sql = generate_sql
    If Trim(sql) <> "" Then
        If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
            sql = InputBox("sql error", "", sql)
            Exit Sub
        End If
    
        Set Me.DataGrid1.DataSource = rs
        Call info(1, "Jumlah data=" & RecordCount(rs), Me.StatusBar1)
    
    
        If Left(Me.cb_jenis.Text, 1) = "4" Then
            p = MsgBox("ubah data NIK?", vbYesNo)
            If p = vbNo Then Exit Sub
            'dari data yang ada, get NIK yang depannya bukan '0000
            jRec = RecordCount(rs)
            If jRec <= 0 Then Exit Sub
        
            rs.MoveFirst
            c = 1
            Do While rs.EOF = False
                Call info(1, "update data. run " & c & "/" & jRec & " -- " & _
                            Round((c / jRec) * 100) & "%", Me.StatusBar1)
        
                nama = cek_null(rs(0))
                npwp = cek_null(rs(1))
            
                sql = "select NIK from pph21tahunan2 where Nama = '" & Trim(nama) & _
                    "' and NPWP = '" & Trim(npwp) & "' and left(NIK,2) <> '00'"
                t = cari_data1(cnn, sql)
                If Trim(t) = "" Then
                    sql = "select NIK from pph21tahunan2 where Nama = '" & Trim(nama) & _
                            "' and NPWP = '" & Trim(npwp) & "'"
                    t = cari_data1(cnn, sql)
                End If
                NIK = t
            
                'update
                sql = "update pph21tahunan2 set NIK = '" & Trim(t) & "' where Nama = '" & Trim(nama) & _
                        "' and NPWP = '" & Trim(npwp) & "'"
                If ExecSQL1(cnn, sql) <> 0 Then
                    sql = InputBox("sql error", "", sql)
                    Exit Do
                End If
            
                rs.MoveNext
                c = c + 1
            Loop
        End If
    Else
        If Left(Me.cb_jenis.Text, 1) = "5" Then
            Call pesan2("penyesuaian P --> TK", 1, vbYellow)
            sql = "update pph21tahunan2 " & _
                    "set PTKP = 'TK' " & _
                    "where P_L = 'P' and PTKP <> 'TK'"
            If ExecSQL1(cnn, sql) <> 0 Then
                sql = InputBox("sql error", "", sql)
                Exit Sub
            End If
            
            Call pesan2("2.Gaji + insentif = 0, hapus", 1, vbYellow)
            sql = "delete " & _
                    "From pph21tahunan2 " & _
                    "Where Gaji + Insentif + THR <= 0"
            If ExecSQL1(cnn, sql) <> 0 Then
                sql = InputBox("sql error", "", sql)
                Exit Sub
            End If
            
            Call pesan2("update NIK", 1, vbYellow)
            
            sql = "Update pph21tahunan2 " & _
                    "set nik =  replace(replace(nik,',',''),'E+15','0000') " & _
                    "where (left(nik,2) = '3,'  or left(nik,2) = '1,') " & _
                    "and length(nik) <= 17"
            If ExecSQL1(cnn, sql) <> 0 Then
                sql = InputBox("sql error", "", sql)
                Exit Sub
            End If
        End If
    End If
    
    
    
    
End Sub




Private Sub cmd_export_Click()
    Dim jRec As Long
    Dim judul As String
    
    Me.disable_Form
    jRec = RecordCount(rs)
    If jRec > 0 Then
        Call create_xls2(rs, "", "", "")
    End If
    Me.Enable_Form
End Sub



Private Sub cmd_Load_Click()
    Me.disable_Form
    Call load_grid
    Me.Enable_Form
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
    
  'set combo
  Me.cb_jenis.Clear
  Me.cb_jenis.AddItem "1. NIK double"
  Me.cb_jenis.AddItem "2. NPWP/NIK/KPP lebih dari 12"
  Me.cb_jenis.AddItem "3. NPWP/NIK/KPP kurang dari 12"
  Me.cb_jenis.AddItem "4. Penyesuaian NIK"
  Me.cb_jenis.AddItem "5. Penyesuaian Data"
  Me.cb_jenis.AddItem "6. NPWP/NIK/KPP kurang dari 12 - per Divisi"
End Sub
