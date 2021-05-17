VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_csv_Ssp 
   ClientHeight    =   7590
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
   ScaleHeight     =   7590
   ScaleWidth      =   12300
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4455
      Left            =   240
      TabIndex        =   17
      Top             =   2760
      Width           =   11895
      Begin VB.CommandButton cmd_export 
         Cancel          =   -1  'True
         Caption         =   "Export CSV"
         Height          =   375
         Left            =   10200
         TabIndex        =   9
         Top             =   3960
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3615
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   11655
         _ExtentX        =   20558
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
      Caption         =   "Load"
      Height          =   375
      Left            =   10560
      TabIndex        =   7
      Top             =   2280
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
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   12015
      Begin VB.OptionButton opt1 
         Caption         =   "Format 1"
         Height          =   375
         Left            =   9600
         TabIndex        =   20
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton opt2 
         Caption         =   "Format 2 (KPP + Divisi)"
         Height          =   375
         Left            =   9600
         TabIndex        =   19
         Top             =   960
         Width           =   2300
      End
      Begin VB.ComboBox cb_pph_ssp 
         Height          =   330
         Left            =   1200
         TabIndex        =   3
         Text            =   "x"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.ComboBox cb_pembetulan 
         Height          =   330
         Left            =   7800
         TabIndex        =   6
         Text            =   "x"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox cb_masa 
         Height          =   330
         Left            =   7800
         TabIndex        =   5
         Text            =   "x"
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cb_tahun 
         Height          =   330
         Left            =   7800
         TabIndex        =   4
         Text            =   "x"
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cb_KPP 
         Height          =   330
         Left            =   1200
         TabIndex        =   2
         Text            =   "Combo1"
         ToolTipText     =   "F2 untuk Filter"
         Top             =   720
         Width           =   5535
      End
      Begin VB.ComboBox cb_divisi 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Text            =   "x"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Format Output"
         Height          =   210
         Left            =   9600
         TabIndex        =   21
         Top             =   360
         Width           =   1020
      End
      Begin VB.Line Line2 
         X1              =   9480
         X2              =   9480
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jenis PPh SSP"
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Pembetulan"
         Height          =   210
         Left            =   6960
         TabIndex        =   16
         Top             =   1140
         Width           =   825
      End
      Begin VB.Line Line1 
         X1              =   6840
         X2              =   6840
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Masa"
         Height          =   210
         Left            =   6960
         TabIndex        =   15
         Top             =   780
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   6960
         TabIndex        =   14
         Top             =   420
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "KPP"
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   780
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Divisi"
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   420
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   7335
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
      Caption         =   "CSV SSP PPh"
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
Attribute VB_Name = "frm_csv_Ssp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsGrid As ADODB.Recordset

Function cek_Isian() As Boolean
    Dim pesan1 As String, t As String
    Dim hasil As Boolean
    
    pesan1 = ""
    hasil = True
    
    'cek divisi
    If Trim(Me.cb_divisi.text) = "" Then
        hasil = False
        pesan1 = pesan1 & "Divisi tidak valid"
    End If
        
    'cek KPP
    If Trim(Me.cb_kpp.text) = "" Then
        hasil = False
        pesan1 = pesan1 & vbCr & "KPP tidak valid"
    End If
    
    If Trim(pesan1) = "" Then
    Else
        MsgBox pesan1
    End If
    
    cek_Isian = hasil
End Function


Private Sub cb_kpp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 And Shift = 0 Then Call load_KPP(Me.cb_kpp, True)
End Sub




Sub disable_Form()
    Me.Frame1.Enabled = False
    Me.Frame2.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame1.Enabled = True
    Me.Frame2.Enabled = True
End Sub


Private Sub cmd_export_Click()
    Dim nmFile As String
    Dim jenisPPh As String, nmPPh As String
    
    nmPPh = "sspPPh"
    
    nmFile = App.Path & "\exp\" & getTimeStamp(Now) & "_" & nmPPh & "_" & get_kode_combo(Me.cb_kpp, "#") & "_" & Trim(Me.cb_tahun) & _
            "_" & Trim(Me.cb_masa) & ".csv"


    Call create_csv(rsGrid, nmFile, , False, "", "")
End Sub

Private Sub cmd_proses_Click()
    Dim jenisPPh As String
    
    On Error GoTo er1
    
    Me.disable_Form
    
    If cek_Isian() = False Then
        Me.Enable_Form
        Exit Sub
    End If
        
        Call load_data_Csv("ssp_pph", 3, 11, _
                                "Kode Form;Masa Pajak SSP;Tahun Pajak SSP;Kode Pembetulan SSP;NTPN (Nomor Transaksi Penerimaan Negara);Tanggal Setor SSP;Jumlah SSP;Kode KAP;Kode Jenis Setoran;Jenis Pajak", _
                                Me.StatusBar1, Me.cb_pph_ssp.text)
    
    Me.Enable_Form
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
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
  
  
  'load combo
  Call load_Divisi(Me.cb_divisi, False, 1, True)
  Call load_KPP(Me.cb_kpp, False, 1)
  Call load_jenisPPhSsp(Me.cb_pph_ssp)
  
  Call load_Tahun2(Me.cb_tahun, "ssp_pph")
  Call load_Masa2(Me.cb_masa, "ssp_pph")
  Call load_Pembetulan2(Me.cb_pembetulan, "ssp_pph")
  Me.opt1.Value = True
  
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
  
  '--- ukuran window
  Me.Width = 12390
  Me.Height = 8025

  
End Sub

Sub load_data_Csv(nmTabel1 As String, kolomAwal As Integer, kolomAkhir As Integer, header1 As String, _
                ByRef sb1 As StatusBar, jenisPPhSSP As String)
    'k01 s/d k15
    
    
    'build rs, dengan kolom k01 s/d kxx
    'di baris pertama, inputkan nama kolom
    'di baris berikutnya, load dari tabel..
    
    Dim c As Integer
    Dim klm1 As String, sql As String
    Dim klm2
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c1 As Long
    
    '---- format tambahan
    If Me.opt1.Value = True Then
    ElseIf Me.opt2.Value = True Then
        kolomAkhir = kolomAkhir + 3
        header1 = header1 & ";npwpKpp;divisi"
    End If
    '====================
    
    'build rs
    klm1 = ""
    For c = 1 To (kolomAkhir - kolomAwal) + 1
        klm1 = klm1 & "k" & adddigit(CLng(c), 2) & ";"
    Next
    klm1 = Left(klm1, Len(klm1) - 1)
    Call create_rs2(rsGrid, klm1)
    
    '--header
    If Trim(header1) <> "" Then
        klm2 = Split(header1, ";")
        rsGrid.AddNew
        For c = 1 To (kolomAkhir - kolomAwal) + 1
            If UBound(klm2) >= c - 1 Then
                rsGrid.Fields(c - 1).Value = klm2(c - 1)
            End If
        Next
        rsGrid.Update
    End If
    
    '-- load data
    sql = create_SQL_PPH(nmTabel1, get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), Me.cb_tahun.text, _
                                Me.cb_masa.text, Me.cb_pembetulan.text, jenisPPhSSP)
    
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    Else
        jRec = RecordCount2(rs)
        If jRec > 0 Then
            rs.MoveFirst
            c1 = 1
            Do While rs.EOF = False
                Call info(2, "Fetch csv. Run " & c1 & "/" & jRec, sb1)
                rsGrid.AddNew
                For c = kolomAwal To kolomAkhir
                    If rs.Fields.Count >= c And c <= rs.Fields.Count - 1 Then
                        rsGrid.Fields(c - kolomAwal) = cek_null(rs.Fields(c))
                    End If
                Next
                
                '--- format2 - npwpkpp + divisi
                    If Me.opt2.Value = True Then
                        If nmTabel1 = "ssp_pph" Then
                            rsGrid.Fields(rsGrid.Fields.Count - 3).Value = cek_null(rs.Fields(1))
                            rsGrid.Fields(rsGrid.Fields.Count - 2).Value = cek_null(rs.Fields(0))
                            rsGrid.Fields(rsGrid.Fields.Count - 1).Value = cek_null(rs.Fields(2))
                        End If
                    End If
                
                rsGrid.Update
                rs.MoveNext
                c1 = c1 + 1
            Loop
        End If
    End If
        
    '----
    
    Set Me.DataGrid1.DataSource = rsGrid
    
End Sub


Function create_SQL_PPH(nmTabel1 As String, npwp_kpp As String, kodeDivisi As String, tahunPajak As String, _
                        masaPajak As String, Pembetulan As String, jenisPPhSSP As String) As String
    
    Dim sql As String, kondisi As String

    sql = "select * from " & nmTabel1
    kondisi = ""
    If Trim(npwp_kpp) = "ALL" Then
    Else
        kondisi = kondisi & " NPWP_KPP = '" & Trim(npwp_kpp) & "'"
    End If
    
    If Trim(kodeDivisi) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " kode_divisi = '" & Trim(kodeDivisi) & "'"
    End If
    
    If Trim(tahunPajak) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Tahun_Pajak = '" & Trim(tahunPajak) & "'"
    End If
    
    If Trim(masaPajak) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Masa_Pajak = '" & Trim(masaPajak) & "'"
    End If
    
    If Trim(Pembetulan) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Pembetulan = '" & Trim(Pembetulan) & "'"
    End If
    
    
    'hanya untuk ssp_pph
        If Trim(jenisPPhSSP) = "" Or Trim(jenisPPhSSP) = "ALL" Then
        Else
            If Trim(kondisi) = "" Then
            Else
                kondisi = kondisi & " AND "
            End If
            kondisi = kondisi & " Jenis_Pajak = '" & Trim(jenisPPhSSP) & "'"
        End If
    
    If Trim(kondisi) = "" Then
    Else
        sql = sql & " WHERE " & kondisi
    End If
    
    
    create_SQL_PPH = sql
End Function

Private Sub Form_Resize()
    If Me.Width - 495 > 0 Then Me.Frame2.Width = Me.Width - 495
    If Me.Height - 3570 > 0 Then Me.Frame2.Height = Me.Height - 3570
    
    If Me.Width - 735 > 0 Then Me.DataGrid1.Width = Me.Width - 735
    If Me.Height - 4410 > 0 Then Me.DataGrid1.Height = Me.Height - 4410
    
    If Me.Height - 4065 > 0 Then Me.cmd_export.Top = Me.Height - 4065
    
End Sub
