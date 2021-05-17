VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Rekap_PPh21Tahunan2 
   ClientHeight    =   7245
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
   ScaleHeight     =   7245
   ScaleWidth      =   12300
   Begin VB.CommandButton cmd_Stop 
      BackColor       =   &H008080FF&
      Caption         =   "Stop Load"
      Height          =   375
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Height          =   5295
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   12015
      Begin VB.CommandButton cmd_expSPT 
         BackColor       =   &H00C0FFC0&
         Caption         =   "export template 1721_bp_A1"
         Height          =   375
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4800
         Width           =   2535
      End
      Begin VB.CommandButton cmd_hitung 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Hitung REKAP"
         Height          =   375
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4793
         Width           =   1455
      End
      Begin VB.CommandButton cmd_info 
         BackColor       =   &H0080FFFF&
         Caption         =   "?"
         Height          =   375
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4793
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmd_export 
         Caption         =   "&Export XLS"
         Height          =   375
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4793
         Width           =   975
      End
      Begin VB.TextBox txt_cari 
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Text            =   "Text1"
         ToolTipText     =   "input dan ENTER"
         Top             =   4793
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   7858
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
         TabIndex        =   14
         Top             =   4875
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " KPP "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   11895
      Begin VB.TextBox txt_divisi 
         Height          =   375
         Left            =   9000
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txt_tahun 
         Height          =   375
         Left            =   6600
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_Load 
         Caption         =   "4. &Load"
         Height          =   375
         Left            =   10560
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cb_kpp 
         Height          =   330
         Left            =   840
         TabIndex        =   1
         Text            =   "x"
         Top             =   382
         Width           =   4695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "3. Filter Divisi "
         Height          =   210
         Left            =   7920
         TabIndex        =   16
         Top             =   442
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2. Tahun"
         Height          =   210
         Left            =   5880
         TabIndex        =   15
         Top             =   442
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "1. KPP"
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   442
         Width           =   465
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   6990
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
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
      Caption         =   "Rekap PPh 21 Tahunan(2)"
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
Attribute VB_Name = "frm_Rekap_PPh21Tahunan2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsGrid As ADODB.Recordset
Dim stopLoad As Boolean

Sub disable_Form()
    Me.Frame2.Enabled = False
    Me.Frame3.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame2.Enabled = True
    Me.Frame3.Enabled = True
End Sub

Sub load_grid(npwp_kpp As String, tahun As String)
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long
    Dim adaError As Boolean
    Dim p, mod1 As Long
        
    '--- ini untuk hitung BUKTI POTONG
    sql = "select distinct '' as bNo1, '' as bBulan, Tahun, " & _
            "'' as bNPWP_KPP, '' as bkdPROYEK, '' as bkdCENTER, " & _
            "Nama, NPWP, NIK, " & _
            "'' as bAlamat, '' as bJabatan, '' as bP_L, " & _
            "'' as bPTKP, '' as bGaji, '' as bTnj_PPh, " & _
            "'' as bTunjangan_Lain, '' as bJHT_JPN, '' as bBruto, " & _
            "'' as bInsentif, '' as bTHR, '' as bLainnya, " & _
            "'' as bPensiun_Potongan_Lain, '' as bid1 from pph21tahunan2 "
       
    If Trim(Me.txt_cari) <> "" Then
        sql = sql & " where (Nama like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NPWP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NIK like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Alamat like '%" & Trim(Me.txt_cari.text) & "%' ) and tahun = '" & tahun & "'"
    Else
        sql = sql & "where tahun = '" & tahun & "'"
    End If
    sql = sql & " order by Nama, NPWP, NIK"
    
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error sql", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
    
    p = MsgBox("RESET ? (Hitung ulang rekap data " & Me.txt_Tahun & ")?", vbYesNo)
    If p = vbYes Then
        '-- delete temp
        sql = "delete from buktipotong where tahun = '" & tahun & "'"
        If ExecSQL1(cnn, sql) <> 0 Then
            sql = InputBox("sql error", "", sql)
            Exit Sub
        End If
        '---
    End If
    
    rs.MoveFirst
    c = 1
    Do While rs.EOF = False
        If stopLoad = True Then Exit Do
        Call info(1, "Run " & c & "/" & jRec & " -- " & Round((c / jRec) * 100, 2) & "% -- Generate Bukti Potong", Me.StatusBar1)
        
        mod1 = c Mod 200
        If mod1 = 0 Then Call dbMySQL_open
        
        Call fetch_Bukti_Potong(Me.txt_Tahun.text, rs, cnn, Me.StatusBar1, adaError)
        
        If adaError = True Then Exit Do
        
        c = c + 1
        rs.MoveNext
    Loop
    
    If Trim(Me.txt_divisi.text) <> "" Then
        Me.Frame3.Caption = Me.Frame3.Caption & " FILTER DIVISI " & Me.txt_divisi.text & " -- "
    End If
    
End Sub


Sub load_grid_lama(npwp_kpp As String, tahun As String)
  Dim a As Integer, c As Integer, jRec As Long
  Dim t As String, sql As String, keteranganFilter As String
  Dim rs As ADODB.Recordset
  Dim nama As String, npwp As String, NIK As String, ptkp As String
  Dim pkp_Setahun As Currency
  
  Dim totalBruto As Currency, totalBiayaJabatan As Currency, totalIuranPensiun As Currency
  Dim totalNetto As Currency, nilaiPtkp As Currency, pph21Setahun As Currency
  
  On Error GoTo er1
  
  Call dbMySQL_open
  
  
  If Trim(tahun) = "" Then
    MsgBox "Tahun tidak valid", vbCritical
    Exit Sub
  End If
  
    '-- referensi
    '0: distinct '' as No1, NPWP_KPP, Nama, " & _
    '3: "NPWP, NIK, Alamat, " & _
    '6: "Jabatan, P_L, PTKP, " & _
    '9; "'' as total_Bruto, '' as total_Netto, '' as nilai_PTKP,  " & _
    '12: "'' as PKP_setahun, '' as PPH21_setahun, '' as pph21_gaji, " & _
    '15: "'' as pph_Gaji_bonus, '' as totalBiayaJabatan, '' as totalIuranPensiun
    '18: '' as jmlData
    '----

  
    sql = "select distinct '' as No1, NPWP_KPP, Nama, " & _
            "NPWP, NIK, Alamat, " & _
            "'' as Jabatan, P_L, PTKP, " & _
            "'' as total_Bruto, '' as total_Netto, '' as nilai_PTKP,  " & _
            "'' as PKP_setahun, '' as PPH21_setahun, '' as pph21_gaji, " & _
            "'' as pph_Gaji_bonus, '' as totalBiayaJabatan, '' as totalIuranPensiun, " & _
            "'' as jmlData " & _
            "from pph21tahunan2 "
    If Trim(npwp_kpp) = "" Or Trim(npwp_kpp) = "ALL" Then
        sql = sql & "where Tahun = '" & Trim(tahun) & "' "
    Else
        sql = sql & "where NPWP_KPP = '" & Trim(npwp_kpp) & "' and Tahun = '" & Trim(tahun) & "' "
    End If
    
    If Trim(Me.txt_divisi.text) <> "" Then
        sql = sql & " and kdCENTER = '" & CekPetik(Me.txt_divisi) & "' "
    End If
    
    If Trim(Me.txt_cari) <> "" Then
        sql = sql & " and (Nama like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NPWP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NIK like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Alamat like '%" & Trim(Me.txt_cari.text) & "%' )"
    End If
    
    sql = sql & " order by Nama, NPWP, NIK"
    
    'sql = InputBox("", "", sql)
    
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockPessimistic, adUseClient) <> 0 Then
        sql = InputBox("Error SQL", "", sql)
        Me.Enable_Form
    Else
        If createRS_duplicate(rs, rsGrid) = True Then
            jRec = RecordCount(rs)
            If jRec > 0 Then
                'copykan isi rs ke rsGrid
                rs.MoveFirst
                c = 1
                Do While rs.EOF = False
                    Call info(1, "Load Grid | Copy | Run " & c & "/" & jRec, Me.StatusBar1)
                    rsGrid.AddNew
                    For a = 0 To rs.Fields.Count - 1
                        rsGrid.Fields(a) = cek_null(rs(a))
                    Next
                    
                    rsGrid.Update
                    c = c + 1
                    rs.MoveNext
                Loop
                
                'manipulasi data
                rsGrid.MoveFirst
                c = 1
                stopLoad = False
                
                Me.cmd_Stop.Visible = True
                Do While rsGrid.EOF = False
                    DoEvents
                    If stopLoad = True Then Exit Do
                    Call info(1, "Load HB Data | cek | Run " & c & "/" & jRec & " -- " & _
                                    Round((c / jRec) * 100, 2) & "%", Me.StatusBar1)
                    nama = cek_null(rsGrid(2))
                    npwp = cek_null(rsGrid(3))
                    NIK = cek_null(rsGrid(4))
                    ptkp = cek_null(rsGrid(8))
                    
                    rsGrid(0) = c
                    rsGrid(6) = tbPph21Tahunan2_getJabatan(npwp, NIK, nama, Me.txt_Tahun)
                    
                    totalNetto = tbPph21Tahunan2_getTotalNetto(totalBruto, totalBiayaJabatan, _
                                totalIuranPensiun, npwp, NIK, nama, tahun)
                    rsGrid(10) = totalNetto
                
                    rsGrid(9) = totalBruto
                    
                    nilaiPtkp = tbM_Ptkp_getNilai(ptkp)
                    rsGrid(11) = nilaiPtkp
                    
                    pkp_Setahun = totalNetto - nilaiPtkp
                    If pkp_Setahun > 0 Then
                        pkp_Setahun = NearestThousand(pkp_Setahun)
                        rsGrid(12) = pkp_Setahun
                    Else
                        pkp_Setahun = 0
                        rsGrid(12) = "0"
                    End If
                    
                    pph21Setahun = get_pph21Setahun(pkp_Setahun)
                    rsGrid(13) = pph21Setahun
                    rsGrid(14) = Round(get_pph21Setahun(pkp_Setahun) / 12, 0)
                    
                    rsGrid(16) = totalBiayaJabatan
                    rsGrid(17) = totalIuranPensiun
                    rsGrid(18) = tbPph21Tahunan2_getJmlData(npwp, NIK, nama, Me.txt_Tahun, Trim(npwp_kpp))
                    
                    rsGrid.Update
                    c = c + 1
                    rsGrid.MoveNext
                Loop
            End If
            Set Me.DataGrid1.DataSource = rsGrid
            Me.Frame3.Caption = " Data | JmlData: " & _
                                RecordCount(rsGrid)
                                
            Call info(1, "JumlahData: " & RecordCount(rsGrid), Me.StatusBar1)
                                
        Else
            Set Me.DataGrid1.DataSource = Nothing
            Me.Frame3.Caption = "  data | ERROR "
        End If
    End If
      
  
  Me.cmd_Stop.Visible = Not True
  '---- status bar
  keteranganFilter = "KPP " & Me.cb_kpp & ". Tahun " & Me.txt_Tahun
  Me.StatusBar1.Panels(2) = "Load finish :: " & keteranganFilter
  Call format_Grid
  
  If Trim(Me.txt_divisi.text) <> "" Then
    Me.Frame3.Caption = Me.Frame3.Caption & " FILTER DIVISI " & Me.txt_divisi.text & " -- "
  End If
  
  Me.Enable_Form
  Exit Sub
er1:
  MsgBox Err.DESCRIPTION
End Sub

Sub format_Grid()
    
    Dim jRec As Long
    Dim c As Integer
    
    jRec = RecordCount(rsGrid)
    If jRec <= 0 Then Exit Sub
    
    
    '-- referensi
    '0: nomor as NoBuktiPotong, tahun, bulan_awal as awal, " & _
    '3: "bulan_akhir as akhir, npwp_pemotong as NPWPKPP, nama_pemotong as KPP, " & _
    '6: "npwp, NIK, Nama, " & _
    '9: "Alamat, Jenis_kelamin as JK, ptkp, " & _
    '12: "jabatan, no_1 as GajiPensiunTht, no_2 as TunjPPh, " & _
    '15: "no_3 as TunjLain, no_4 as Honor, no_5 as as Premi, " & _
    '18: "no_6 as Lain, no_7 as BonusThr, no_8 as 1sd7, " & _
    '21: "no_9 as jabPensiun, no_10 as IuranPensiun, no_11 as 9sd10, " & _
    '24: "no_12 as net811, no_13 as netPajakSblm, no_14 as netHitungPPh, " & _
    '27: "no_15 as PenghasilanTKP, no_16 as pkpSetahun, no_17 as pphSetahun, " & _
    '30: "no_18 as PPhTelahDiptg, no_19 as pphTerutang, no_20 as PPhTlhDipotong " & _
    '----
        
        For c = 0 To rsGrid.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
            
            'kolom kecil
            If c = 1 Or c = 2 Or c = 3 Or c = 10 Or c = 11 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).Width = 500
            End If
    
            'If c = 23 Then
            '    Me.DataGrid1.Columns(c).Alignment = dbgCenter
            '    Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
            '    Me.DataGrid1.Columns(c).Width = 900
            'End If
    
            If c >= 13 And c <= 32 Then
                
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1300
            End If
        Next

End Sub

Private Sub cb_kpp_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 113 And Shift = 0 Then Call load_KPP(Me.cb_kpp, True)
End Sub

Private Sub cmd_export_Click()
    Dim jRec As Long
    Dim judul As String
    
    judul = "KPP " & Me.cb_kpp & ". Tahun " & Me.txt_Tahun
    If Trim(Me.txt_divisi) = "" Then
    Else
        judul = judul & " Filter Divisi " & Me.txt_divisi
    End If
    
    If Trim(Me.txt_cari) = "" Then
    Else
        judul = judul & " Filter Cari: " & Me.txt_cari
    End If
    
    '-- referensi
    '0: nomor as NoBuktiPotong, tahun, bulan_awal as awal, " & _
    '3: "bulan_akhir as akhir, npwp_pemotong as NPWPKPP, nama_pemotong as KPP, " & _
    '6: "npwp, NIK, Nama, " & _
    '9: "Alamat, Jenis_kelamin as JK, ptkp, " & _
    '12: "jabatan, no_1 as GajiPensiunTht, no_2 as TunjPPh, " & _
    '15: "no_3 as TunjLain, no_4 as Honor, no_5 as as Premi, " & _
    '18: "no_6 as Lain, no_7 as BonusThr, no_8 as 1sd7, " & _
    '21: "no_9 as jabPensiun, no_10 as IuranPensiun, no_11 as 9sd10, " & _
    '24: "no_12 as net811, no_13 as netPajakSblm, no_14 as netHitungPPh, " & _
    '27: "no_15 as PenghasilanTKP, no_16 as pkpSetahun, no_17 as pphSetahun, " & _
    '30: "no_18 as PPhTelahDiptg, no_19 as pphTerutang, no_20 as PPhTlhDipotong " & _
    '----
    
    Me.disable_Form
    jRec = RecordCount(rsGrid)
    If jRec > 0 Then
        Call create_xls2(rsGrid, judul, "13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,31", "")
    End If
    Me.Enable_Form
End Sub

Sub generate_noUrut_BuktiPotong(tahun As String, Vnpwp_kpp As String, _
                                Optional digitAwal As Integer = 1)
    'list per npwp_kpp per nama, npwp, nik, tahun
    
    Dim sql As String
    Dim c As Long, jRec As Long, mod1 As Long
    Dim d1 As Long
    Dim nama As String, npwp As String, NIK As String, npwp_kpp2 As String
    Dim nomor As String, npwp_kpp As String
    Dim rs As ADODB.Recordset
    
    
    '-- set no Urut direksi
    sql = "select distinct npwp_kpp, nama, npwp, nik, tahun " & _
            "From pph21tahunan2 " & _
            "where nama in ('tumiyana','agus purbianto','abdul haris tatang','lukman hidayat'," & _
            "'M. Aprindy', 'mohammad toha fauzi') " & _
            " and tahun = '" & tahun & "' " & _
            "order by npwp_kpp, nama "
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rs)
    If jRec > 0 Then
    
        c = 1
        d1 = digitAwal
        rs.MoveFirst
        npwp_kpp2 = cek_null(rs(0))
        Do While rs.EOF = False
            Call info(2, "set No DIREKSI. Run " & c & "/" & jRec & " -- " & Round((c / jRec) * 100, 0) & "%", _
                    frMenu1.StatusBar1)
        
            npwp_kpp = cek_null(rs(0))
            If npwp_kpp <> npwp_kpp2 Then d1 = 1
            npwp_kpp2 = npwp_kpp
            nama = cek_null(rs(1))
            npwp = cek_null(rs(2))
            NIK = cek_null(rs(3))
            nomor = "1.1.12" & Right(tahun, 2) & adddigit(d1, 7)
        
            'update
            sql = "update pph21tahunan2 set no_urut_buktipotong = '" & nomor & "' where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
                "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "' and NPWP_KPP = '" & npwp_kpp & "'"
            If ExecSQL1(cnn, sql) <> 0 Then
                sql = InputBox("sql error", "", sql)
                Exit Do
            End If
    
            rs.MoveNext
            c = c + 1
            d1 = d1 + 1
        Loop
    End If
    '---------------------------------------------------------------------
    
    If Trim(Vnpwp_kpp) = "" Or Vnpwp_kpp = "ALL" Then
        sql = "select distinct npwp_kpp, nama, npwp, nik, tahun " & _
            "From pph21tahunan2 " & _
            "where nama not in ('tumiyana','agus purbianto','abdul haris tatang','lukman hidayat'," & _
            "'M. Aprindy', 'mohammad toha fauzi') " & _
            " and tahun = '" & tahun & "' " & _
            "order by npwp_kpp, nama "
    Else
        sql = "select distinct npwp_kpp, nama, npwp, nik, tahun " & _
            "From pph21tahunan2 " & _
            "where npwp_kpp = '" & Trim(Vnpwp_kpp) & "' " & _
            "and nama not in ('tumiyana','agus purbianto','abdul haris tatang','lukman hidayat'," & _
            "'M. Aprindy', 'mohammad toha fauzi') " & _
            " and tahun = '" & tahun & "' " & _
            "order by npwp_kpp, nama "
    End If
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
    
    c = 1
    rs.MoveFirst
    npwp_kpp2 = cek_null(rs(0))
    '-- ini KPP jakarta
    If npwp_kpp2 = "010016137093000" Then
        If digitAwal > 7 Then
            d1 = digitAwal
        Else
            d1 = digitAwal + 7
        End If
    Else
        d1 = digitAwal
    End If
    Do While rs.EOF = False
        Call info(2, "set No. Run " & c & "/" & jRec & " -- " & Round((c / jRec) * 100, 2) & "% - NoBuktiPotong", _
                    frMenu1.StatusBar1)
        
        mod1 = c Mod 1000
        If mod1 = 0 Then Call dbMySQL_open
        
        npwp_kpp = cek_null(rs(0))
        If npwp_kpp <> npwp_kpp2 Then d1 = digitAwal
        npwp_kpp2 = npwp_kpp
        nama = cek_null(rs(1))
        npwp = cek_null(rs(2))
        NIK = cek_null(rs(3))
        nomor = "1.1.12" & Right(tahun, 2) & adddigit(d1, 7)
        
        'update
        sql = "update pph21tahunan2 set no_urut_buktipotong = '" & nomor & "' where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "' and NPWP_KPP = '" & npwp_kpp & "'"
        If ExecSQL1(cnn, sql) <> 0 Then
            sql = InputBox("sql error", "", sql)
            Exit Do
        End If
    
        rs.MoveNext
        c = c + 1
        d1 = d1 + 1
    Loop
    Call pesan2("Proses set no Urut tahun " & Me.txt_Tahun.text & " selesai", 1)
End Sub

Private Sub cmd_expSPT_Click()
    Dim sql As String, namaFile As String
    Dim rs As ADODB.Recordset
    
    sql = "select bulan_akhir as 'Masa Pajak', tahun as 'Tahun Pajak', '0' as Pembetulan, " & _
            "nomor as 'Nomor Bukti Potong', bulan_awal as 'Masa Perolehan Awal', bulan_akhir as 'Masa Perolehan Akhir', " & _
            "NPWP, NIK, Nama, " & _
            "Alamat, jenis_kelamin as 'Jenis Kelamin', " & _
            "get_status_ptkp(ptkp) as 'Status PTKP', get_jml_tanggungan(ptkp) as 'Jumlah Tanggungan', jabatan as 'Nama Jabatan', " & _
            "'N' as 'WP Luar Negeri', '' as 'Kode Negara', '21-100-01' as 'Kode Pajak', " & _
            "no_1 as 'Jumlah 1', no_2 as 'Jumlah 2', no_3 as 'Jumlah 3', " & _
            "no_4 as 'Jumlah 4', no_5 as 'Jumlah 5', no_6 as 'Jumlah 6', " & _
            "no_7 as 'Jumlah 7', no_8 as 'Jumlah 8', no_9 as 'Jumlah 9', " & _
            "no_10 as 'Jumlah 10', no_11 'Jumlah 11', no_12 as 'Jumlah 12', " & _
            "no_13 as 'Jumlah 13', no_14 as 'Jumlah 14', no_15 as 'Jumlah 15', " & _
            "no_16 as 'Jumlah 16', no_17 as 'Jumlah 17', no_18 as 'Jumlah 18', " & _
            "no_19 as 'Jumlah 19', no_20 as 'Jumlah 20', '' as 'Status Pindah', " & _
            "replace(npwp_ttd,'.','') as 'NPWP Pemotong', nama_ttd as 'Nama Pemotong', concat('28/',bulan_akhir,'/',tahun) as 'Tanggal Bukti Potong' " & _
            "From buktipotong "
    If Trim(Me.cb_kpp) = "ALL" Or Trim(Me.cb_kpp) = "" Then
    Else
        sql = sql & " where npwp_pemotong = '" & get_kode_combo(Me.cb_kpp, "#") & "'"
    End If
    
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error sql", "", sql)
        Exit Sub
    End If
    namaFile = InputBox("nama file", "", App.Path & "\exp\1721_bp_A1.csv")
    Call create_csv(rs, namaFile)
End Sub

Private Sub cmd_hitung_Click()
    Dim t1 As Date, t2 As Date
    Dim sql As String
    Dim p
    Dim digitAwal As Integer
    
    On Error GoTo er1
    Me.disable_Form
    t1 = Now
    
    
    p = MsgBox("Update NOMOR bukti potong?", vbYesNo)
    If p = vbYes Then
        digitAwal = CInt(InputBox("Digit Awal No", "", "2001"))
        Call generate_noUrut_BuktiPotong(Year(Now) - 1, get_kode_combo(Me.cb_kpp, "#"), _
                                         digitAwal)
    End If
    
    
    p = MsgBox("Proses Hitung Rekap Data?", vbYesNo)
    If p = vbYes Then
        stopLoad = False
        Me.cmd_Stop.Visible = True
        Call load_grid(get_kode_combo(Me.cb_kpp, "#"), Me.txt_Tahun.text)
    End If
    
    
    p = MsgBox("Hitung Beban per Divisi?", vbYesNo)
    If p = vbYes Then
        Call hitung_beban_Divisi
    End If
    
    t2 = Now
    MsgBox "Finish in " & CDate(CDate(t2) - CDate(t1)), , vbInformation
    
    Me.Enable_Form
    Exit Sub
er1:
    Me.Enable_Form
End Sub

Sub hitung_beban_Divisi()
    'report biaya beban per divisi, hitung dari tabel bukti potong
    '- list dari data buktipotong
    'Per nama, npwp, nik, tahun..
    'cari bulan paling akhir, cari pphAkhir & jumlah_data
    '- nilaibeban = pphAkhir / jumlahdata
    '- update di pph21tahunan2 : kolom bebanpph
    
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim c As Long, jRec As Long, mod1 As Long
    Dim jmlData As Integer, jmlPPh As Currency
    Dim nama As String, NIK As String, npwp As String, tahun As String
    Dim nilaiBeban As Currency
    
    sql = "select nama, npwp, nik, tahun, sum(no_20) from buktipotong " & _
            "where tahun = '" & Trim(Me.txt_Tahun) & "' " & _
            "group by nama, npwp, nik, tahun " & _
            "order by nama, npwp, nik"
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error sql", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rs)
    If jRec <= 1 Then Exit Sub
    c = 1
    Do While rs.EOF = False
        Call info(1, "Run Hitung Beban. " & c & "/" & jRec & " -- " & Round((c / jRec) * 100, 2) & "%", _
                    Me.StatusBar1)
        mod1 = c Mod 800
        If mod1 = 0 Then Call dbMySQL_open
        nama = cek_null(rs(0))
        npwp = cek_null(rs(1))
        NIK = cek_null(rs(2))
        tahun = cek_null(rs(3))
        
        If nama = "SURYANA SAEFUL ROHMAN" Then
            mod1 = 0
        End If
        
        'cari jumlah data
        sql = "select count(*) " & _
                "from pph21tahunan2 where nik = '" & Trim(NIK) & "' and nama = '" & Trim(nama) & _
                "' and npwp = '" & Trim(npwp) & "' and tahun = '" & Trim(tahun) & "'"
        jmlData = cari_data1(cnn, sql, True)
        '----
        jmlPPh = cek_Money(rs(4))
        
        If jmlData <= 0 Then
            nilaiBeban = 0
        Else
            If jmlPPh <= 0 Then
                nilaiBeban = 0
            Else
                nilaiBeban = Round(jmlPPh / jmlData, 0)
            End If
        End If
        
        
        
        
        'update
        sql = "update pph21tahunan2 set nilai_beban = '" & nilaiBeban & "' where nik = '" & Trim(NIK) & "' and nama = '" & Trim(nama) & _
                "' and npwp = '" & Trim(npwp) & "' and tahun = '" & Trim(tahun) & "'"
        If ExecSQL1(cnn, sql) <> 0 Then
            sql = InputBox("error run", "", sql)
            Exit Do
        End If
        rs.MoveNext
        c = c + 1
    Loop
End Sub

Private Sub cmd_info_Click()
    Dim t As String
    
    t = "BRUTO = Gaji + Tnj_PPh + JHT_JPN + Tunjangan_Lain" & vbCr & _
        "Netto = totalBruto -  totalBiaya Jabatan  - totalIuran Pensiun/Potongan Lain" & vbCr & _
        "BiayaJabatan = if(0.05 * Bruto<500000;0.05 * Bruto<500000;500000)"
    
    
    MsgBox t, vbInformation
End Sub

Private Sub cmd_load_Click()
    Dim t1 As Date, t2 As Date
    Dim sql As String, t As String
    Dim kondisi As String, jumlahData As Integer
    
    On Error GoTo er1
    Me.disable_Form
    t1 = Now
    stopLoad = False
    Me.cmd_Stop.Visible = True
    
    If Trim(Me.txt_Tahun) = "" Then
        MsgBox "Tahun harus terisi", vbInformation
        Exit Sub
    End If
    
    'tampilkan grid
    kondisi = ""
    sql = "select nomor as NoBuktiPotong, tahun, bulan_awal as awal, " & _
            "bulan_akhir as akhir, npwp_pemotong as NPWPKPP, nama_pemotong as KPP, " & _
            "npwp, NIK, Nama, " & _
            "Alamat, Jenis_kelamin as JK, ptkp, " & _
            "jabatan, no_1 as GajiPensiunTht, no_2 as TunjPPh, " & _
            "no_3 as TunjLain, no_4 as Honor, no_5 as Premi, " & _
            "no_6 as Lain, no_7 as BonusThr, no_8 as 1sd7, " & _
            "no_9 as jabPensiun, no_10 as IuranPensiun, no_11 as 9sd10, " & _
            "no_12 as net811, no_13 as netPajakSblm, no_14 as netHitungPPh, " & _
            "no_15 as PenghasilanTKP, no_16 as pkpSetahun, no_17 as pphSetahun, " & _
            "no_18 as PPhTelahDiptg, no_19 as pphTerutang, no_20 as PPhTlhDipotong, " & _
            "kdCENTER " & _
            "from buktipotong "
    If Trim(Me.cb_kpp) = "ALL" Or Trim(Me.cb_kpp) = "" Then
    Else
        kondisi = kondisi & "npwp_pemotong = '" & get_kode_combo(Me.cb_kpp, "#") & "' "
    End If
    
    If Trim(Me.txt_divisi) = "" Then
    Else
        If Trim(Me.cb_kpp) = "ALL" Or Trim(Me.cb_kpp) = "" Then
            kondisi = kondisi & " kdCENTER = '" & Trim(Me.txt_divisi) & "' "
        Else
            kondisi = kondisi & "and kdCENTER = '" & Trim(Me.txt_divisi) & "' "
        End If
    End If
    
    If Trim(Me.txt_cari) <> "" Then
        If (Trim(Me.cb_kpp) = "ALL" Or Trim(Me.cb_kpp) = "") And Trim(Me.txt_divisi) = "" Then
            kondisi = kondisi & "  (Nama like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NPWP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NIK like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Alamat like '%" & Trim(Me.txt_cari.text) & "%' )"
        Else
            kondisi = kondisi & " and (Nama like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NPWP like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "NIK like '%" & Trim(Me.txt_cari.text) & "%' or " & _
                    "Alamat like '%" & Trim(Me.txt_cari.text) & "%' )"
        End If
    End If
    
    'tahun
        If Trim(kondisi) = "" Then
            kondisi = kondisi & "  tahun = '" & Trim(Me.txt_Tahun) & "'"
        Else
            kondisi = kondisi & " and tahun = '" & Trim(Me.txt_Tahun) & "'"
        End If
    '------------
    If Trim(kondisi) <> "" Then
        sql = sql & " where " & kondisi
    End If
    sql = sql & " order by Nama, NIK, npwp "
    
    t = ""
        
    Do While IsNumeric(t) = False
        t = InputBox("Jumlah data yang akan ditampilkan ? (0:semua data)", "", "0")
        jumlahData = CInt(t)
    Loop
    
    If jumlahData = 0 Then
    Else
        sql = sql & " limit " & jumlahData
    End If
    
    If OpenRecordSet(cnn, rsGrid, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error sql", "", sql)
        Exit Sub
    End If
            
    Set Me.DataGrid1.DataSource = rsGrid
    Call format_Grid
    Call info(2, "Jumlah data = " & RecordCount(rsGrid), Me.StatusBar1)
    
    t2 = Now
    MsgBox "Finish in " & CDate(CDate(t2) - CDate(t1)), , vbInformation
    Me.cmd_Stop.Visible = False
    Me.Enable_Form
    Exit Sub
er1:
    Me.Enable_Form
End Sub

Private Sub cmd_Stop_Click()
    stopLoad = True
    Me.cmd_Stop.Visible = False
End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
    
  'ukuran awal
  Me.Width = 12390
  Me.Height = 7680
  '-----------
  
  Me.txt_Tahun.text = Year(Now) - 1
  Me.txt_cari.text = ""
  Me.txt_divisi.text = ""
  Me.cmd_Stop.Visible = False
    
'---------------------
    If dbMySQL_open = True Then
    Else
        MsgBox "Open Database Tidak Sukses", vbExclamation
    End If
    '---------------------
    
  'load combo
  Call load_KPP(Me.cb_kpp, False, 1)
  
  'get level
  '2 : operator gedung
  '3 : UKP
  
  '---------
  
  Level1 = tbPengguna_getLevel1(frMenu1.nmLogin)
  Call info(2, "Level: " & Level1, Me.StatusBar1)
  
  MsgBox "Lakukan hitung rekap lengkap sebelum menampilkan data rekap / bukti potong." & vbCr & _
        "Lakukan reset data HANYA jika ada perubahan data record", vbInformation
  
End Sub


Private Sub Form_Resize()
    Me.Shape1.Width = Me.Width
    Me.lb_caption.Width = Me.Width
    
    If Me.Width - 375 > 0 Then Me.Frame3.Width = Me.Width - 375
    If Me.Height - 2385 > 0 Then Me.Frame3.Height = Me.Height - 2385
    
    If Me.Height - 2887 > 0 Then Me.txt_cari.Top = Me.Height - 2887
    Me.Label6.Top = Me.txt_cari.Top
    Me.cmd_info.Top = Me.txt_cari.Top
    Me.cmd_expSPT.Top = Me.txt_cari.Top
    Me.cmd_hitung.Top = Me.txt_cari.Top
    Me.cmd_export.Top = Me.txt_cari.Top
    
    If Me.Height - 1080 > 0 Then Me.cmd_Stop.Top = Me.Height - 1080

    If Me.Width - 615 > 0 Then Me.DataGrid1.Width = Me.Width - 615
    If Me.Height - 3225 > 0 Then Me.DataGrid1.Height = Me.Height - 3225
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call dbMySQL_close
End Sub

Private Sub txt_cari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmd_load_Click
    End If
End Sub
