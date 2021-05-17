VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frm_repPPh 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3405
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
   ScaleHeight     =   3405
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   " 2. Jenis Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   7335
      Begin VB.OptionButton opt_buktipotong 
         Caption         =   "Bukti Potong"
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton opt_rekap 
         Caption         =   "Rekap"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton opt_detil 
         Caption         =   "Detil"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport CR 
      Left            =   7320
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmd_proses 
      Cancel          =   -1  'True
      Caption         =   "Print"
      Height          =   495
      Left            =   10320
      TabIndex        =   9
      Top             =   2520
      Width           =   1815
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
      Begin VB.ComboBox cb_proyek 
         Height          =   330
         Left            =   4680
         TabIndex        =   3
         Text            =   "x"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cb_masa 
         Height          =   330
         Left            =   8160
         TabIndex        =   6
         Text            =   "x"
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cb_tahun 
         Height          =   330
         Left            =   8160
         TabIndex        =   5
         Text            =   "x"
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cb_KPP 
         Height          =   330
         Left            =   1200
         TabIndex        =   4
         Text            =   "Combo1"
         ToolTipText     =   "F2 untuk Filter"
         Top             =   1080
         Width           =   5535
      End
      Begin VB.ComboBox cb_jenisPajak 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cb_divisi 
         Height          =   330
         Left            =   1200
         TabIndex        =   2
         Text            =   "x"
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Proyek"
         Height          =   210
         Left            =   4080
         TabIndex        =   18
         Top             =   780
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   7320
         X2              =   7320
         Y1              =   240
         Y2              =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Masa"
         Height          =   210
         Left            =   7560
         TabIndex        =   16
         Top             =   780
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   7560
         TabIndex        =   15
         Top             =   420
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "KPP"
         Height          =   210
         Left            =   240
         TabIndex        =   14
         Top             =   1140
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pajak "
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   420
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Divisi"
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   780
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3150
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
   Begin VB.Label lb_caption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Report SPT PPh"
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
Attribute VB_Name = "frm_repPPh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim doFetchRecord As Boolean

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
    
    'cek jenispajak
    t = get_kode_combo(Me.cb_jenisPajak, ".")
    If t = "1" Or t = "2" Or t = "3" Or t = "4" Or t = "5" Or t = "6" Or _
        t = "7" Or t = "8" Or t = "9" Or t = "10" Or t = "11" Or t = "12" Then
    Else
        hasil = False
        pesan1 = pesan1 & vbCr & "Jenis Pajak tidak valid"
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

Private Sub cb_divisi_Click()
    Dim jenisPPh As String
    
    Me.disable_Form
    doFetchRecord = True
    jenisPPh = get_kode_combo(Me.cb_jenisPajak, ".")
    If Trim(jenisPPh) = "1" Then
        Call load_Proyek(get_kode_combo(Me.cb_divisi, "-"), "pph15", Me.cb_proyek)
    ElseIf Trim(jenisPPh) = "2" Then
        Call load_Proyek(get_kode_combo(Me.cb_divisi, "-"), "pph23", Me.cb_proyek)
    ElseIf Trim(jenisPPh) = "3" Then
        Call load_Proyek(get_kode_combo(Me.cb_divisi, "-"), "pph21tf", Me.cb_proyek)
    ElseIf Trim(jenisPPh) = "4" Then
        Call load_Proyek(get_kode_combo(Me.cb_divisi, "-"), "pph21bulanan", Me.cb_proyek)
    ElseIf Trim(jenisPPh) = "5" Then
        Call load_Proyek(get_kode_combo(Me.cb_divisi, "-"), "pph21tahunan", Me.cb_proyek)
    ElseIf Trim(jenisPPh) = "6" Then
        Call load_Proyek(get_kode_combo(Me.cb_divisi, "-"), "pph22", Me.cb_proyek)
    ElseIf Trim(jenisPPh) = "7" Then
        Call load_Proyek(get_kode_combo(Me.cb_divisi, "-"), "pph26", Me.cb_proyek)
    ElseIf Trim(jenisPPh) = "8" Then
        Call load_Proyek(get_kode_combo(Me.cb_divisi, "-"), "pph42_konstruksi", Me.cb_proyek)
    ElseIf Trim(jenisPPh) = "9" Then
        Call load_Proyek(get_kode_combo(Me.cb_divisi, "-"), "pph42_sewa", Me.cb_proyek)
    ElseIf Trim(jenisPPh) = "10" Then
        Call load_Proyek(get_kode_combo(Me.cb_divisi, "-"), "pph42_obligasi", Me.cb_proyek)
    End If
    Me.Enable_Form
End Sub

Private Sub cb_jenisPajak_Click()
    Dim jenisPPh As String
    Dim Level1 As Integer
    
    Me.disable_Form
    doFetchRecord = True
    Call dbMySQL_open
    
    jenisPPh = get_kode_combo(Me.cb_jenisPajak, ".")
    If Trim(jenisPPh) = "1" Then
        Call load_Tahun_pph15(Me.cb_tahun)
        Call load_Masa_pph15(Me.cb_masa)
    ElseIf Trim(jenisPPh) = "2" Then
        Call load_Tahun2(Me.cb_tahun, "pph23")
        Call load_Masa2(Me.cb_masa, "pph23")
    ElseIf Trim(jenisPPh) = "3" Then
        Call load_Tahun2(Me.cb_tahun, "pph21tf")
        Call load_Masa2(Me.cb_masa, "pph21tf")
    ElseIf Trim(jenisPPh) = "4" Then
        Call load_Tahun2(Me.cb_tahun, "pph21bulanan")
        Call load_Masa2(Me.cb_masa, "pph21bulanan")
    ElseIf Trim(jenisPPh) = "5" Then
        Call load_Tahun2(Me.cb_tahun, "pph21tahunan")
        Call load_Masa2(Me.cb_masa, "pph21tahunan")
    ElseIf Trim(jenisPPh) = "6" Then
        Call load_Tahun2(Me.cb_tahun, "pph22")
        Call load_Masa2(Me.cb_masa, "pph22")
    ElseIf Trim(jenisPPh) = "7" Then
        Call load_Tahun2(Me.cb_tahun, "pph26")
        Call load_Masa2(Me.cb_masa, "pph26")
    ElseIf Trim(jenisPPh) = "8" Then
        Call load_Tahun2(Me.cb_tahun, "pph42_konstruksi")
        Call load_Masa2(Me.cb_masa, "pph42_konstruksi")
    ElseIf Trim(jenisPPh) = "9" Then
        Call load_Tahun2(Me.cb_tahun, "pph42_sewa")
        Call load_Masa2(Me.cb_masa, "pph42_sewa")
    ElseIf Trim(jenisPPh) = "10" Then
        Call load_Tahun2(Me.cb_tahun, "pph42_obligasi")
        Call load_Masa2(Me.cb_masa, "pph42_obligasi")
    ElseIf Trim(jenisPPh) = "11" Then
        Call load_Tahun2(Me.cb_tahun, "pph21_bwhptkp")
        Call load_Masa2(Me.cb_masa, "pph21_bwhptkp")
    ElseIf Trim(jenisPPh) = "12" Then
        Call load_Tahun2(Me.cb_tahun, "pph21pesangon")
        Call load_Masa2(Me.cb_masa, "pph21pesangon")
    Else
        Me.cb_tahun.Clear
        Me.cb_masa.Clear
    End If
    
    Level1 = tbPengguna_getLevel1(frMenu1.nmLogin)
    If Level1 = 2 Then
        Call cb_divisi_Click
    End If
    
    Me.Enable_Form
End Sub

Private Sub cb_kpp_KeyDown(KeyCode As Integer, Shift As Integer)
    doFetchRecord = True
    If KeyCode = 113 And Shift = 0 Then Call load_KPP(Me.cb_kpp, True)
    
End Sub




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



Sub fetch_dbRep_PPh15(npwp_kpp As String, kodeDivisi As String, kodeProyek As String, tahunPajak As String, masaPajak As String, _
                        ByRef sb1 As StatusBar)
                        
    Dim cnnTemp As ADODB.connection
    Dim sql As String, kondisi As String
    Dim rsSumber As ADODB.Recordset, rsTujuan As ADODB.Recordset
    Dim jRec As Long, c As Long, a As Integer
    
    'open cnnTemp
    Call db_Access_open(cnnTemp, App.Path & "\data\dbrep.mdb")
    
    'delete
    sql = "delete from pph15"
    If ExecSQL1(cnnTemp, sql) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    
    'insert
    sql = "select * from pph15 "
    kondisi = ""
    If Trim(npwp_kpp) = "ALL" Then
    Else
        kondisi = kondisi & " npwp_kpp = '" & Trim(npwp_kpp) & "'"
    End If
    
    If Trim(kodeDivisi) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " kode_divisi = '" & Trim(kodeDivisi) & "'"
    End If
    
    If IsNumeric(kodeProyek) = True Then
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " kd_proyek = '" & Trim(kodeProyek) & "'"
    End If
    
    
    
    If Trim(tahunPajak) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " tahun_pajak = '" & Trim(tahunPajak) & "'"
    End If
    
    If Trim(masaPajak) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " masa_pajak = '" & Trim(masaPajak) & "'"
    End If
    
    If Trim(kondisi) = "" Then
    Else
        sql = sql & " WHERE " & kondisi
    End If
    '---
    
    'MsgBox sql
    
    If OpenRecordSet(cnn, rsSumber, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    Else
        jRec = RecordCount2(rsSumber)
        If jRec > 0 Then
            '-- open target
            sql = "select * from pph15"
            If OpenRecordSet(cnnTemp, rsTujuan, sql, adOpenDynamic, adLockPessimistic, adUseClient) <> 0 Then
                sql = InputBox("error", "", sql)
                Exit Sub
            End If
            '--
        
            rsSumber.MoveFirst
            c = 1
            Do While rsSumber.EOF = False
                Call info(2, "Fetch divisi. Run " & c & "/" & jRec, sb1)
                
                rsTujuan.AddNew
                'For a = 0 To rsSumber.Fields.Count - 1
                For a = 0 To rsSumber.Fields.Count - 2
                    rsTujuan.Fields(a).Value = cek_null(rsSumber.Fields(a).Value)
                Next
                rsTujuan.Update
                
                rsSumber.MoveNext
                c = c + 1
            Loop
            Set rsSumber = Nothing
        End If
    End If
    
End Sub

Sub proses_bukti_potong()
    Dim jenisPPh As String
    jenisPPh = get_kode_combo(Me.cb_jenisPajak, ".")
        
    If dbMySQL_open = False Then
        Me.Enable_Form
        Exit Sub
    End If
    
    If jenisPPh = "2" Then
        If doFetchRecord = True Then
          Call fetch_dbRep_PPhX2(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), _
                                Me.cb_proyek.text, Me.cb_tahun.text, Me.cb_masa.text, _
                                Me.StatusBar1, "pph23")
        End If
        Call tampil_report(CR, App.Path & "\rep\rep_BP_pph23.rpt", 85)
        
        'Dim Appl As New CRAXDRT.Application
        'Dim Report As New CRAXDRT.Report
        'Set Report = Appl.OpenReport(App.Path & "\rep\rep_BP_pph23.rpt")
        'Report.EnableParameterPrompting = False

        'With Report
        '    .ExportOptions.FormatType = crEFTPortableDocFormat
        '    .ExportOptions.DestinationType = crEDTDiskFile
        '    .ExportOptions.DiskFileName = App.Path + "\exp\test.pdf"
            ' location & the file name

        '    .ExportOptions.PDFExportAllPages = True
        '    .Export (False)
        'End With
        
    ElseIf jenisPPh = "6" Then
        If doFetchRecord = True Then
          Call fetch_dbRep_PPhX2(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), _
                                Me.cb_proyek.text, Me.cb_tahun.text, Me.cb_masa.text, _
                                Me.StatusBar1, "pph22")
        End If
        Call tampil_report(CR, App.Path & "\rep\rep_BP_pph22.rpt", 85)
        
    ElseIf jenisPPh = "8" Then
        If doFetchRecord = True Then
          Call fetch_dbRep_PPhX2(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), _
                                Me.cb_proyek.text, Me.cb_tahun.text, Me.cb_masa.text, _
                                Me.StatusBar1, "pph42_konstruksi")
        End If
        Call tampil_report(CR, App.Path & "\rep\rep_BP_pph42_konstruksi.rpt", 85)
    ElseIf jenisPPh = "9" Then
        If doFetchRecord = True Then
          Call fetch_dbRep_PPhX2(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), _
                                Me.cb_proyek.text, Me.cb_tahun.text, Me.cb_masa.text, _
                                Me.StatusBar1, "pph42_sewa")
        End If
        Call tampil_report(CR, App.Path & "\rep\rep_BP_pph42_sewa.rpt", 85)
    ElseIf jenisPPh = "10" Then
        If doFetchRecord = True Then
          Call fetch_dbRep_PPhX2(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), _
                                Me.cb_proyek.text, Me.cb_tahun.text, Me.cb_masa.text, _
                                Me.StatusBar1, "pph42_obligasi")
        End If
        Call tampil_report(CR, App.Path & "\rep\rep_BP_pph42_obligasi.rpt", 85)
    Else
        Call pesan2("do nothing..")
    End If
    
    doFetchRecord = False
    Me.Enable_Form
End Sub



Private Sub cb_masa_Click()
    doFetchRecord = True
End Sub

Private Sub cb_tahun_Click()
    doFetchRecord = True
End Sub

Private Sub cmd_proses_Click()
    Dim jenisPPh As String
    
    On Error GoTo er1
    Me.disable_Form
    
    If cek_Isian() = False Then
        Me.Enable_Form
        Exit Sub
    End If
    
    
    If dbMySQL_open = False Then
        Me.Enable_Form
        Exit Sub
    End If
    
    Call create_ds_Access("c:\dbpph2.dsn", App.Path & "\data\", App.Path & "\data\dbrep.mdb")
    Call create_ds_Access("c:\dbpph.dsn", App.Path & "\data\", App.Path & "\data\dbrep.mdb")
    
    Call fetch_dbRep_Divisi(get_kode_combo(Me.cb_divisi, "-"), Me.StatusBar1)
    Call fetch_dbRep_KPP(get_kode_combo(Me.cb_kpp, "#"), Me.StatusBar1)
    
    
    
    If Me.opt_buktipotong.Value = True Then
        Call proses_bukti_potong
        Exit Sub
    End If
    
    jenisPPh = get_kode_combo(Me.cb_jenisPajak, ".")
    If Trim(jenisPPh) = "1" Then
        'MsgBox "a"
        If doFetchRecord = True Then
            Call fetch_dbRep_PPh15(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), Me.cb_proyek.text, Me.cb_tahun.text, _
                                Me.cb_masa.text, Me.StatusBar1)
        End If
        'MsgBox "b"
        If Me.opt_detil.Value = True Then
            Call tampil_report(CR, App.Path & "\rep\repPph15.rpt", 85)
        Else
            Call tampil_report(CR, App.Path & "\rep\repPph15_rekap.rpt", 85)
        End If
        'MsgBox "c"
    ElseIf Trim(jenisPPh) = "2" Then
        If doFetchRecord = True Then
            Call fetch_dbRep_PPhX(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), Me.cb_proyek.text, Me.cb_tahun.text, _
                                Me.cb_masa.text, Me.StatusBar1, "pph23")
        End If
        If Me.opt_detil.Value = True Then
            Call tampil_report(CR, App.Path & "\rep\repPph23.rpt", 85)
        Else
            Call tampil_report(CR, App.Path & "\rep\repPph23_rekap.rpt", 85)
        End If
    ElseIf Trim(jenisPPh) = "3" Then
        If doFetchRecord = True Then
            Call fetch_dbRep_PPhX(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), Me.cb_proyek.text, Me.cb_tahun.text, _
                                Me.cb_masa.text, Me.StatusBar1, "pph21tf")
        End If
        If Me.opt_detil.Value = True Then
            Call tampil_report(CR, App.Path & "\rep\repPph21TF.rpt", 85)
        Else
            Call tampil_report(CR, App.Path & "\rep\repPph21TF_rekap.rpt", 85)
        End If
    ElseIf Trim(jenisPPh) = "4" Then
        If doFetchRecord = True Then
            Call fetch_dbRep_PPhX(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), Me.cb_proyek.text, Me.cb_tahun.text, _
                                Me.cb_masa.text, Me.StatusBar1, "pph21bulanan")
        End If
        
        If Me.opt_detil.Value = True Then
            Call tampil_report(CR, App.Path & "\rep\repPph21Bulanan.rpt", 85)
        Else
            Call tampil_report(CR, App.Path & "\rep\repPph21Bulanan_rekap.rpt", 85)
        End If
    ElseIf Trim(jenisPPh) = "5" Then
        If doFetchRecord = True Then
            Call fetch_dbRep_PPhX(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), Me.cb_proyek.text, Me.cb_tahun.text, _
                                Me.cb_masa.text, Me.StatusBar1, "pph21tahunan")
        End If
        
        If Me.opt_detil.Value = True Then
            Call tampil_report(CR, App.Path & "\rep\repPph21Tahunan.rpt", 85)
        Else
            Call tampil_report(CR, App.Path & "\rep\repPph21Tahunan_rekap.rpt", 85)
        End If
    ElseIf Trim(jenisPPh) = "6" Then
        If doFetchRecord = True Then
            Call fetch_dbRep_PPhX(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), Me.cb_proyek.text, Me.cb_tahun.text, _
                                Me.cb_masa.text, Me.StatusBar1, "pph22")
        End If
        
        If Me.opt_detil.Value = True Then
            Call tampil_report(CR, App.Path & "\rep\repPph22.rpt", 85)
        Else
            Call tampil_report(CR, App.Path & "\rep\repPph22_rekap.rpt", 85)
        End If
    ElseIf Trim(jenisPPh) = "7" Then
        If doFetchRecord = True Then
            Call fetch_dbRep_PPhX(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), Me.cb_proyek.text, Me.cb_tahun.text, _
                                Me.cb_masa.text, Me.StatusBar1, "pph26")
        End If
        If Me.opt_detil.Value = True Then
            Call tampil_report(CR, App.Path & "\rep\repPph26.rpt", 85)
        Else
            Call tampil_report(CR, App.Path & "\rep\repPph26_rekap.rpt", 85)
        End If
    ElseIf Trim(jenisPPh) = "8" Then
        If doFetchRecord = True Then
            Call fetch_dbRep_PPhX(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), Me.cb_proyek.text, Me.cb_tahun.text, _
                                Me.cb_masa.text, Me.StatusBar1, "pph42_konstruksi")
        End If
        If Me.opt_detil.Value = True Then
            Call tampil_report(CR, App.Path & "\rep\repPph42_konstruksi.rpt", 85)
        Else
            Call tampil_report(CR, App.Path & "\rep\repPph42_konstruksi_rekap.rpt", 85)
        End If
    ElseIf Trim(jenisPPh) = "9" Then
        If doFetchRecord = True Then
            Call fetch_dbRep_PPhX(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), Me.cb_proyek.text, Me.cb_tahun.text, _
                                Me.cb_masa.text, Me.StatusBar1, "pph42_sewa")
        End If
        If Me.opt_detil.Value = True Then
            Call tampil_report(CR, App.Path & "\rep\repPph42_sewa.rpt", 85)
        Else
            Call tampil_report(CR, App.Path & "\rep\repPph42_sewa_rekap.rpt", 85)
        End If
    ElseIf Trim(jenisPPh) = "10" Then
        If doFetchRecord = True Then
            Call fetch_dbRep_PPhX(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), Me.cb_proyek.text, Me.cb_tahun.text, _
                                Me.cb_masa.text, Me.StatusBar1, "pph42_obligasi")
        End If
        If Me.opt_detil.Value = True Then
            Call tampil_report(CR, App.Path & "\rep\repPph42_obligasi.rpt", 85)
        Else
            Call tampil_report(CR, App.Path & "\rep\repPph42_obligasi_rekap.rpt", 85)
        End If
    ElseIf Trim(jenisPPh) = "11" Then
        If doFetchRecord = True Then
            Call fetch_dbRep_PPhX(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), Me.cb_proyek.text, Me.cb_tahun.text, _
                                Me.cb_masa.text, Me.StatusBar1, "pph21_bwhptkp")
        End If
        If Me.opt_detil.Value = True Then
            Call tampil_report(CR, App.Path & "\rep\repPph21bwhptkp.rpt", 85)
        Else
            Call tampil_report(CR, App.Path & "\rep\repPph21bwhptkp_rekap.rpt", 85)
        End If
    ElseIf Trim(jenisPPh) = "12" Then
        If doFetchRecord = True Then
            Call fetch_dbRep_PPhX(get_kode_combo(Me.cb_kpp, "#"), get_kode_combo(Me.cb_divisi, "-"), Me.cb_proyek.text, Me.cb_tahun.text, _
                                Me.cb_masa.text, Me.StatusBar1, "pph21pesangon")
        End If
        If Me.opt_detil.Value = True Then
            Call tampil_report(CR, App.Path & "\rep\repPph21pesangon.rpt", 85)
        Else
            Call tampil_report(CR, App.Path & "\rep\repPph21pesangon_rekap.rpt", 85)
        End If
    Else
        Call pesan2("No Reports", , vbYellow)
    End If
    doFetchRecord = False
    
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
  Call load_jenisPPh(Me.cb_jenisPajak)
  Call load_Divisi(Me.cb_divisi, False, 1, True)
  Call load_KPP(Me.cb_kpp, False, 1)
  
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
  
    doFetchRecord = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Kill "C:\dbpph.dsn"
    Call dbMySQL_close
End Sub


Private Sub opt_buktipotong_Click()
    doFetchRecord = True
End Sub

Private Sub opt_detil_Click()
    doFetchRecord = True
End Sub
