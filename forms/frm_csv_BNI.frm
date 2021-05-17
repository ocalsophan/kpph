VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_csv_BNI 
   ClientHeight    =   7380
   ClientLeft      =   300
   ClientTop       =   1110
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
   Begin VB.CommandButton cmd_info 
      BackColor       =   &H00C0FFC0&
      Caption         =   "?"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "info akun FTP"
      Top             =   1920
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4455
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   12015
      Begin VB.CommandButton cmd_xls 
         Caption         =   "Export XLS"
         Height          =   375
         Left            =   8880
         TabIndex        =   15
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cmd_export 
         Caption         =   "Export CSV"
         Height          =   375
         Left            =   10200
         TabIndex        =   4
         ToolTipText     =   "Export / Upload"
         Top             =   3960
         Width           =   1575
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
      TabIndex        =   6
      Top             =   600
      Width           =   12015
      Begin VB.ListBox List1 
         Height          =   900
         Left            =   6600
         TabIndex        =   14
         Top             =   240
         Width           =   5295
      End
      Begin VB.TextBox txt_masa 
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txt_Tahun 
         Height          =   375
         Left            =   4560
         TabIndex        =   11
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
         TabIndex        =   9
         Top             =   780
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   4080
         TabIndex        =   8
         Top             =   420
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Divisi"
         Height          =   210
         Left            =   240
         TabIndex        =   7
         Top             =   420
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
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
      Caption         =   "CSV Mandiri MFT"
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
   Begin VB.Menu frMenu 
      Caption         =   "Menu"
      Begin VB.Menu mntaxinquiry 
         Caption         =   "tax_inquiry_report"
      End
      Begin VB.Menu mnImpTaxInqury 
         Caption         =   "Import tax inqury report"
      End
   End
End
Attribute VB_Name = "frm_csv_BNI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsGrid As ADODB.Recordset
Dim rsData As ADODB.Recordset
Dim sudah_ada_header As Boolean
Dim nmFile As String


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


Private Sub cb_divisi_Click()
    'Call isi_list
End Sub

Private Sub cmd_export_Click()
    Dim nmFile_simpan As String, nmfile2 As String, sql As String, fileWinscp As String
    Dim jenisPPh As String, nmPPh As String
    Dim p, s
    Dim cek As String, res As String
    Dim d1 As Date, d2 As Date, isecords As Long, i As Long, jRec As Long
    'Payment Agustus 2019_MPNG2SSP_20190808_175427
    
    '-- cek apakah file winscp sudah ada ?
    fileWinscp = "c:\Program Files (x86)\WinSCP\WinSCP.com"
    If Dir(fileWinscp) <> "" Then
        'MsgBox "File exists"
    Else
        MsgBox "WINSCP File does not exist. Please Install First"
        Exit Sub
    End If
    '------
    
    If sudah_ada_header = False Then
        Call create_header
    Else
        nmFile_simpan = App.Path & "\exp\" & nmFile & ".txt"
        nmfile2 = nmFile & ".txt"
        Call create_csv2(rsGrid, nmFile_simpan, tbVariabel_get("delimiter"), False)
        res = tbVariabel_get("waktu_upload")
        p = MsgBox("Upload File ?" & vbCr & _
                    "Last time upload: " & res & vbCr & _
                    "Jika ada konfirmasi jaringan, pilih 'Allow'", vbYesNo)
        If p = vbYes Then
                 
            cek = tbBniDirect_get(get_kode_combo(Me.cb_divisi, "-"), "user1")
            If Trim(cek) = "" Then
                Call pesan2("user name invalid")
            Else
                'kasih durasi 30 detik
                If res = "" Then
                    d1 = Now
                Else
                    d1 = res
                End If
                d2 = Now
                isecords = DateDiff("s", d1, d2)
                
                If isecords < 60 Then
                    Do While (i <= (60 - isecords) * 300000)
                        DoEvents
                        'Sleep 100
                        i = i + 1
                    Loop
                End If
                
                
                Call UploadFile(tbVariabel_get("ftpip"), _
                                cek, _
                                tbBniDirect_get(get_kode_combo(Me.cb_divisi, "-"), "pass1"), _
                               tbVariabel_get("ftpport"), nmfile2)
                Call tbVariabel_set("waktu_upload", Now)
                
                s = MsgBox("Pengiriman Sukses ?", vbYesNo)
                If s = vbYes Then
                    'baca dari rsdata, masukkan ke tabel log
                    jRec = RecordCount(rsData)
                    If jRec > 0 Then
                        rsData.MoveFirst
                        Do While rsData.EOF = False
                            sql = "select F_insert_log_export_mandiri('" & _
                                cek_null(rsData(0)) & "','" & cek_null(rsData(1)) & "','" & _
                                cek_null(rsData(2)) & "','" & cek_null(rsData(3)) & "','" & _
                                cek_null(rsData(4)) & "','" & cek_null(rsData(5)) & "','" & _
                                cek_null(rsData(6)) & "','" & cek_null(rsData(7)) & "','" & _
                                cek_null(rsData(9)) & "','" & cek_null(rsData(11)) & "','" & _
                                cek_null(rsData(12)) & "','" & cek_null(rsData(13)) & "','" & _
                                cek_null(rsData(14)) & "','" & cek_null(rsData(15)) & "','" & _
                                get_kode_combo(Me.cb_divisi, "-") & "','" & nmfile2 & "')"
                            'sql = InputBox("", "", sql)
                            Call ExecSQL1(cnn, sql)
                            
                            rsData.MoveNext
                        Loop
                    
                    End If
                End If
            End If
        End If
    End If
    Call set_tombol_load
End Sub

Private Sub cmd_info_Click()
    Dim t As String
    Dim r As Long
    
    
    t = "FTP Address : 182.253.5.26" & vbCr & _
        "uname: pp" & vbCr & _
        "pass: PpPajak2019" & vbCr & _
        "port : 21" & _
        "Wajib install winscp, install di C:\"
    Call Shell("explorer.exe https://winscp.net/eng/download.php?TBiframe", vbNormalFocus)
    'url1 = "explorer.exe " & "http://acc.ptpp.co.id/kpph"
    'File1 = InputBox("", "", File1)
    'Call Shell(url1, vbNormalFocus)
        
    MsgBox t, vbInformation
End Sub

Sub isi_list()
    Dim im As String, t As String, sql As String
    Dim run_in_the_coming_days As Integer
    Dim counter1 As String
    
    Me.disable_Form
    Me.List1.Clear
    'nama file, norekening, jml data, jumlah total, uname,pwd : xxx, email, instruction mode,
    'transaction_instruction_date, run_in_the_coming_days, session
    
    counter1 = ""
    nmFile = tbVariabel_get("company_id") & "16" & Format(Now, "ddmmyyyy") & _
            Left(get_kode_combo(Me.cb_divisi, "-"), 3) & counter1
    
    'counter, cek jika nama file tsb sudah ada di db
    sql = "select max(mid(nmfile,21,3)) from log_export_mandiri where left(nmfile,20) = '" & Left(nmFile, 20) & "'"
    t = cari_data1(cnn, sql, True)
    counter1 = adddigit(CInt(t) + 1, 3)
    nmFile = tbVariabel_get("company_id") & "16" & Format(Now, "ddmmyyyy") & _
            Left(get_kode_combo(Me.cb_divisi, "-"), 3) & counter1
    
    
    
    Call setListInfo(Me.List1, "FileName:" & nmFile)
    Call setListInfo(Me.List1, "Division No:" & get_kode_combo(Me.cb_divisi, "-"))
    Call setListInfo(Me.List1, "AcctNo:" & tbBniDirect_getNorek(get_kode_combo(Me.cb_divisi, "-")))
    
    Call setListInfo(Me.List1, "Record Count:" & RecordCount(rsData))
    
    Call setListInfo(Me.List1, "Payment Amount:" & Format(total_nominal_data, "###,###"))
    Call setListInfo(Me.List1, "uname:" & tbBniDirect_get(get_kode_combo(Me.cb_divisi, "-"), "user1"))
    Call setListInfo(Me.List1, "email:" & tbVariabel_get("email"))
    
    im = tbVariabel_get("instruction_mode")
    If Trim(im) = "1" Then
        Call setListInfo(Me.List1, "instruction_mode:" & im)
    Else
        run_in_the_coming_days = cek_Int(tbVariabel_get("run_in_the_coming_days"))
        Call setListInfo(Me.List1, "instruction_mode:" & im)
        Call setListInfo(Me.List1, "run_in_the_coming_days:" & tbVariabel_get("run_in_the_coming_days"))
        Call setListInfo(Me.List1, "transaction_instruction_date:" & Format(DateAdd("d", run_in_the_coming_days, Now), "yyyymmdd"))
        Call setListInfo(Me.List1, "session:" & tbVariabel_get("session"))
    End If
    Call setListInfo(Me.List1, "delimiter:" & tbVariabel_get("delimiter"))
    Me.Enable_Form
End Sub

Function total_nominal_data() As Currency
    Dim totalNominal As Currency
    Dim c As Long, jRec As Long
    
    totalNominal = 0
    c = 1
    jRec = RecordCount(rsData)
    If jRec > 0 Then
    rsData.MoveFirst
        Do While rsData.EOF = False
            Call info_progress(Me.StatusBar1, 2, c, jRec, "Cek total")
            totalNominal = totalNominal + cek_Money(rsData(12))
            c = c + 1
            rsData.MoveNext
        Loop
    End If
    total_nominal_data = totalNominal
End Function

Private Sub cmd_proses_Click()
    On Error GoTo er1
    'create dulu grid isinya..
    'kemudian 2 baris atasnya
    
    Dim kdDivisi As String, jmlBaris As Integer
    Dim c As Long, jRec As Long
    
    Me.disable_Form
    kdDivisi = get_kode_combo(Me.cb_divisi, "-")
    
    'load perjenispajak, utk semua npwpkpp, per masa
    
    Call create_rs2(rsData, "k1;k2;k3;k4;k5;k6;k7;k8;k9;k10;k11;k12;k13;k14;k15;" & _
                            "k16;k17;k18;k19;k20;k21")
    'pph21
    'pph22
    'pph23
    'pph23, jika terisi pph_yang_dipotong__5, kode setor 100,
    ' Lainnya 104
    'pph26
    'pph42_sewa
    'pph42_konstruksi
    'pph 15, 41128, 410
    Call load_dataPajak("pph21bulanan", kdDivisi, Me.txt_Tahun, Me.txt_masa, Me.StatusBar1)
    Call load_dataPajak("pph21tf", kdDivisi, Me.txt_Tahun, Me.txt_masa, Me.StatusBar1)
    Call load_dataPajak("pph22", kdDivisi, Me.txt_Tahun, Me.txt_masa, Me.StatusBar1)
    Call load_dataPajak("pph23", kdDivisi, Me.txt_Tahun, Me.txt_masa, Me.StatusBar1)
    Call load_dataPajak("pph23_jasa", kdDivisi, Me.txt_Tahun, Me.txt_masa, Me.StatusBar1)
    Call load_dataPajak("pph23_deviden", kdDivisi, Me.txt_Tahun, Me.txt_masa, Me.StatusBar1)
    
    Call load_dataPajak("pph26", kdDivisi, Me.txt_Tahun, Me.txt_masa, Me.StatusBar1)
    Call load_dataPajak("pph26_tarif1", kdDivisi, Me.txt_Tahun, Me.txt_masa, Me.StatusBar1)
    Call load_dataPajak("pph26_tarif2", kdDivisi, Me.txt_Tahun, Me.txt_masa, Me.StatusBar1)
    Call load_dataPajak("pph26_tarif3", kdDivisi, Me.txt_Tahun, Me.txt_masa, Me.StatusBar1)
    Call load_dataPajak("pph26_tarif45", kdDivisi, Me.txt_Tahun, Me.txt_masa, Me.StatusBar1)
    
    Call load_dataPajak("pph42_sewa", kdDivisi, Me.txt_Tahun, Me.txt_masa, Me.StatusBar1)
    Call load_dataPajak("pph42_konstruksi", kdDivisi, Me.txt_Tahun, Me.txt_masa, Me.StatusBar1)
    Call load_dataPajak("pph15", kdDivisi, Me.txt_Tahun, Me.txt_masa, Me.StatusBar1)
    
     
    
    'add header1,
    'Call create_header
    sudah_ada_header = False
    Call isi_list
    Set Me.DataGrid1.DataSource = rsData
    
    Me.Frame2.Caption = " Jumlah Data: " & jmlBaris & ". Total Nominal: " & _
                        Format(total_nominal_data, "###,###")
    Me.Enable_Form
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
    Me.Enable_Form
End Sub

Sub create_header()
    Dim jmlBaris As Integer
    Dim im As String
    Dim run_in_the_coming_days As Integer
    
    jmlBaris = RecordCount(rsData)
    
    If jmlBaris <= 0 Then
        Call pesan2("no data")
        Exit Sub
    End If
    
    Call create_rs2(rsGrid, "k1;k2;k3;k4;k5;k6;k7;k8;k9;k10;k11;k12;k13;k14;k15;" & _
                            "k16;k17;k18;k19;k20;k21")
    rsGrid.AddNew
    rsGrid.Fields(0) = "P"
    rsGrid.Fields(1) = tbBniDirect_getNorek(get_kode_combo(Me.cb_divisi, "-"))
    rsGrid.Fields(2) = jmlBaris
    im = tbVariabel_get("instruction_mode")
    
    If Trim(im) = "1" Then
        rsGrid.Fields(3) = Trim(im)
        rsGrid.Fields(4) = ""
        rsGrid.Fields(5) = ""
    ElseIf Trim(im) = "2" Then
        rsGrid.Fields(3) = Trim(im)
        run_in_the_coming_days = cek_Int(tbVariabel_get("run_in_the_coming_days"))
        rsGrid.Fields(4) = Format(DateAdd("d", run_in_the_coming_days, Now), "YYYYMMDD")
        rsGrid.Fields(5) = tbVariabel_get("session")
    End If
    rsGrid.Update
    'copy data
    If createRS_duplicateContent(rsData, rsGrid, Me.StatusBar1, False) = False Then
        MsgBox "error copy rs", vbCritical
        Exit Sub
    End If
    Set Me.DataGrid1.DataSource = rsGrid
    sudah_ada_header = True
End Sub

Sub load_dataPajak(jenisPajak As String, kdDivisi As String, tahun As String, _
                    masa As String, ByRef sb1 As StatusBar)
    
    Dim sql As String, ket As String, masa2 As String
    Dim rs As ADODB.Recordset
    Dim c As Long, jRec As Long, d As Integer
    Dim cust_Ref As String
    
    masa2 = adddigit(CLng(masa), 2)
    cust_Ref = Left(kdDivisi, 3) & "_" & Right(tahun, 2) & masa2 & "_" & Format(Now, "yymmdd")
    ket = jenisPajak & " " & Format(DateSerial(CInt(tahun), CInt(masa), 1), "MMMM YYYY")
    Call dbMySQL_open
    If jenisPajak = "pph21bulanan" Then
        cust_Ref = cust_Ref & "A"
        sql = "Select pph21bulanan.NPWP_KPP, left(mkpp.nama,30), left(mkpp.alamat,50), " & _
                "f_getkotaKpp(pph21bulanan.NPWP_KPP) as k4, " & _
                "'' as k5, '411121' as k6, '100' as k7, " & _
                "'" & masa2 & "' as k8, '" & masa2 & "' as k9, " & _
                "pph21bulanan.Tahun_Pajak as k10, '' as k11, 'IDR' as k12, " & _
                "Sum(pph21bulanan.Jumlah_PPh) as k13, '" & cust_Ref & "' as k14, " & _
                "f_getVariabel('email') as k15, '" & ket & "' as k16, '' as k17, '' as k18, " & _
                "pph21bulanan.NPWP_KPP as k19, '' as k20, 'E' as k21 " & _
                "From pph21bulanan Left Join mkpp On mkpp.npwp = pph21bulanan.NPWP_KPP " & _
                "Where pph21bulanan.kode_divisi = '" & kdDivisi & "' And " & _
                "pph21bulanan.Masa_Pajak = '" & masa2 & "' And " & _
                "pph21bulanan.Tahun_Pajak = '" & tahun & "' " & _
                "Group By pph21bulanan.kode_divisi, pph21bulanan.NPWP_KPP, mkpp.nama, " & _
                "mkpp.alamat, pph21bulanan.Masa_Pajak, pph21bulanan.Tahun_Pajak"
    ElseIf jenisPajak = "pph21tf" Then
        cust_Ref = cust_Ref & "B"
            sql = "Select " & jenisPajak & ".NPWP_KPP, left(mkpp.nama,30), left(mkpp.alamat,50), " & _
                "f_getkotaKpp(" & jenisPajak & ".NPWP_KPP) as k4, " & _
                "'' as k5, '411121' as k6, '100' as k7, " & _
                "'" & masa2 & "' as k8, '" & masa2 & "' as k9, " & _
                "" & jenisPajak & ".Tahun_Pajak as k10, '' as k11, 'IDR' as k12, " & _
                "Sum(" & jenisPajak & ".Jumlah_PPh) as k13, '" & cust_Ref & "' as k14, " & _
                "f_getVariabel('email') as k15, '" & ket & "' as k16, '' as k17, '' as k18, " & _
                "" & jenisPajak & ".NPWP_KPP as k19, '' as k20, 'E' as k21 " & _
                "From " & jenisPajak & " Left Join mkpp On mkpp.npwp = " & jenisPajak & ".NPWP_KPP " & _
                "Where " & jenisPajak & ".kode_divisi = '" & kdDivisi & "' And " & _
                "" & jenisPajak & ".Masa_Pajak = '" & masa2 & "' And " & _
                "" & jenisPajak & ".Tahun_Pajak = '" & tahun & "' " & _
                "Group By " & jenisPajak & ".kode_divisi, " & jenisPajak & ".NPWP_KPP, mkpp.nama, " & _
                "mkpp.alamat, " & jenisPajak & ".Masa_Pajak, " & jenisPajak & ".Tahun_Pajak"
    ElseIf jenisPajak = "pph22" Then
        cust_Ref = cust_Ref & "C"
        sql = "Select " & jenisPajak & ".NPWP_KPP, left(mkpp.nama,30), left(mkpp.alamat,50), " & _
                "f_getkotaKpp(" & jenisPajak & ".NPWP_KPP) as k4, " & _
                "'' as k5, '411122' as k6, '100' as k7, " & _
                "'" & masa2 & "' as k8, '" & masa2 & "' as k9, " & _
                "" & jenisPajak & ".Tahun_Pajak as k10, '' as k11, 'IDR' as k12, " & _
                "Sum(" & jenisPajak & ".nilai_pph) as k13, '" & cust_Ref & "' as k14, " & _
                "f_getVariabel('email') as k15, '" & ket & "' as k16, '' as k17, '' as k18, " & _
                "" & jenisPajak & ".NPWP_KPP as k19, '' as k20, 'E' as k21 " & _
                "From " & jenisPajak & " Left Join mkpp On mkpp.npwp = " & jenisPajak & ".NPWP_KPP " & _
                "Where " & jenisPajak & ".kode_divisi = '" & kdDivisi & "' And " & _
                "" & jenisPajak & ".Masa_Pajak = '" & masa2 & "' And " & _
                "" & jenisPajak & ".Tahun_Pajak = '" & tahun & "' " & _
                "Group By " & jenisPajak & ".kode_divisi, " & jenisPajak & ".NPWP_KPP, mkpp.nama, " & _
                "mkpp.alamat, " & jenisPajak & ".Masa_Pajak, " & jenisPajak & ".Tahun_Pajak"
    ElseIf jenisPajak = "pph23" Then
        cust_Ref = cust_Ref & "D"
        sql = "Select " & jenisPajak & ".NPWP_KPP, left(mkpp.nama,30), left(mkpp.alamat,50), " & _
                "f_getkotaKpp(" & jenisPajak & ".NPWP_KPP) as k4, " & _
                "'' as k5, '411124' as k6, '100' as k7, " & _
                "'" & masa2 & "' as k8, '" & masa2 & "' as k9, " & _
                "" & jenisPajak & ".Tahun_Pajak as k10, '' as k11, 'IDR' as k12, " & _
                "Sum(" & jenisPajak & ".Jumlah_PPh_Yang_Dipotong) as k13, '" & cust_Ref & "' as k14, " & _
                "f_getVariabel('email') as k15, '" & ket & "' as k16, '' as k17, '' as k18, " & _
                "" & jenisPajak & ".NPWP_KPP as k19, '' as k20, 'E' as k21 " & _
                "From " & jenisPajak & " Left Join mkpp On mkpp.npwp = " & jenisPajak & ".NPWP_KPP " & _
                "Where " & jenisPajak & ".kode_divisi = '" & kdDivisi & "' And " & _
                "" & jenisPajak & ".Masa_Pajak = '" & masa2 & "' And " & _
                "" & jenisPajak & ".Tahun_Pajak = '" & tahun & "' and PPh_Yang_Dipotong__5 > 0 " & _
                "Group By " & jenisPajak & ".kode_divisi, " & jenisPajak & ".NPWP_KPP, mkpp.nama, " & _
                "mkpp.alamat, " & jenisPajak & ".Masa_Pajak, " & jenisPajak & ".Tahun_Pajak"
    ElseIf jenisPajak = "pph23_jasa" Then
        'utk tatif bukan 5 dan bukan 1
        cust_Ref = cust_Ref & "E"
        sql = "Select pph23.NPWP_KPP, left(mkpp.nama,30), left(mkpp.alamat,50), " & _
                "f_getkotaKpp(pph23.NPWP_KPP) as k4, " & _
                "'' as k5, '411124' as k6, '104' as k7, " & _
                "'" & masa2 & "' as k8, '" & masa2 & "' as k9, " & _
                "pph23.Tahun_Pajak as k10, '' as k11, 'IDR' as k12, " & _
                "Sum(pph23.Jumlah_PPh_Yang_Dipotong) as k13, '" & cust_Ref & "' as k14, " & _
                "f_getVariabel('email') as k15, '" & ket & "' as k16, '' as k17, '' as k18, " & _
                "pph23.NPWP_KPP as k19, '' as k20, 'E' as k21 " & _
                "From pph23 Left Join mkpp On mkpp.npwp = pph23.NPWP_KPP " & _
                "Where pph23.kode_divisi = '" & kdDivisi & "' And " & _
                "pph23.Masa_Pajak = '" & masa2 & "' And " & _
                "pph23.Tahun_Pajak = '" & tahun & "' and not(PPh_Yang_Dipotong__5 > 0) " & _
                "and not(PPh_Yang_Dipotong__1 > 0) " & _
                "Group By pph23.kode_divisi, pph23.NPWP_KPP, mkpp.nama, " & _
                "mkpp.alamat, pph23.Masa_Pajak, pph23.Tahun_Pajak"
    ElseIf jenisPajak = "pph23_deviden" Then
        'utk tarif pph 1
        cust_Ref = cust_Ref & "E"
        sql = "Select pph23.NPWP_KPP, left(mkpp.nama,30), left(mkpp.alamat,50), " & _
                "f_getkotaKpp(pph23.NPWP_KPP) as k4, " & _
                "'' as k5, '411124' as k6, '101' as k7, " & _
                "'" & masa2 & "' as k8, '" & masa2 & "' as k9, " & _
                "pph23.Tahun_Pajak as k10, '' as k11, 'IDR' as k12, " & _
                "Sum(pph23.Jumlah_PPh_Yang_Dipotong) as k13, '" & cust_Ref & "' as k14, " & _
                "f_getVariabel('email') as k15, '" & ket & "' as k16, '' as k17, '' as k18, " & _
                "pph23.NPWP_KPP as k19, '' as k20, 'E' as k21 " & _
                "From pph23 Left Join mkpp On mkpp.npwp = pph23.NPWP_KPP " & _
                "Where pph23.kode_divisi = '" & kdDivisi & "' And " & _
                "pph23.Masa_Pajak = '" & masa2 & "' And " & _
                "pph23.Tahun_Pajak = '" & tahun & "' and (PPh_Yang_Dipotong__1 > 0) " & _
                "Group By pph23.kode_divisi, pph23.NPWP_KPP, mkpp.nama, " & _
                "mkpp.alamat, pph23.Masa_Pajak, pph23.Tahun_Pajak"
    ElseIf jenisPajak = "pph26" Then
        cust_Ref = cust_Ref & "F"
        sql = "Select " & jenisPajak & ".NPWP_KPP, left(mkpp.nama,30), left(mkpp.alamat,50), " & _
                "f_getkotaKpp(" & jenisPajak & ".NPWP_KPP) as k4, " & _
                "'' as k5, '411127' as k6, '100' as k7, " & _
                "'" & masa2 & "' as k8, '" & masa2 & "' as k9, " & _
                "" & jenisPajak & ".Tahun_Pajak as k10, '' as k11, 'IDR' as k12, " & _
                "Sum(" & jenisPajak & ".Jumlah_PPh_Yang_Dipotong) as k13, '" & cust_Ref & "' as k14, " & _
                "f_getVariabel('email') as k15, '" & ket & "' as k16, '' as k17, '' as k18, " & _
                "" & jenisPajak & ".NPWP_KPP as k19, '' as k20, 'E' as k21 " & _
                "From " & jenisPajak & " Left Join mkpp On mkpp.npwp = " & jenisPajak & ".NPWP_KPP " & _
                "Where " & jenisPajak & ".kode_divisi = '" & kdDivisi & "' And " & _
                "" & jenisPajak & ".Masa_Pajak = '" & masa2 & "' And " & _
                "" & jenisPajak & ".Tahun_Pajak = '" & tahun & "' " & _
                "and (PPh_Yang_Dipotong__1 + PPh_Yang_Dipotong__2 + PPh_Yang_Dipotong__3 + PPh_Yang_Dipotong__4 + PPh_Yang_Dipotong__5 <= 0) " & _
                "Group By " & jenisPajak & ".kode_divisi, " & jenisPajak & ".NPWP_KPP, mkpp.nama, " & _
                "mkpp.alamat, " & jenisPajak & ".Masa_Pajak, " & jenisPajak & ".Tahun_Pajak"
    ElseIf jenisPajak = "pph26_tarif1" Then
        cust_Ref = cust_Ref & "G"
        sql = "Select pph26.NPWP_KPP, left(mkpp.nama,30), left(mkpp.alamat,50), " & _
                "f_getkotaKpp(pph26.NPWP_KPP) as k4, " & _
                "'' as k5, '411127' as k6, '101' as k7, " & _
                "'" & masa2 & "' as k8, '" & masa2 & "' as k9, " & _
                "pph26.Tahun_Pajak as k10, '' as k11, 'IDR' as k12, " & _
                "Sum(pph26.Jumlah_PPh_Yang_Dipotong) as k13, '" & cust_Ref & "' as k14, " & _
                "f_getVariabel('email') as k15, '" & ket & "' as k16, '' as k17, '' as k18, " & _
                "pph26.NPWP_KPP as k19, '' as k20, 'E' as k21 " & _
                "From pph26 Left Join mkpp On mkpp.npwp = pph26.NPWP_KPP " & _
                "Where pph26.kode_divisi = '" & kdDivisi & "' And " & _
                "pph26.Masa_Pajak = '" & masa2 & "' And " & _
                "pph26.Tahun_Pajak = '" & tahun & "' and PPh_Yang_Dipotong__1 > 0 " & _
                "Group By pph26.kode_divisi, pph26.NPWP_KPP, mkpp.nama, " & _
                "mkpp.alamat, pph26.Masa_Pajak, pph26.Tahun_Pajak"
    ElseIf jenisPajak = "pph26_tarif2" Then
        cust_Ref = cust_Ref & "H"
        sql = "Select pph26.NPWP_KPP, left(mkpp.nama,30), left(mkpp.alamat,50), " & _
                "f_getkotaKpp(pph26.NPWP_KPP) as k4, " & _
                "'' as k5, '411127' as k6, '102' as k7, " & _
                "'" & masa2 & "' as k8, '" & masa2 & "' as k9, " & _
                "pph26.Tahun_Pajak as k10, '' as k11, 'IDR' as k12, " & _
                "Sum(pph26.Jumlah_PPh_Yang_Dipotong) as k13, '" & cust_Ref & "' as k14, " & _
                "f_getVariabel('email') as k15, '" & ket & "' as k16, '' as k17, '' as k18, " & _
                "pph26.NPWP_KPP as k19, '' as k20, 'E' as k21 " & _
                "From pph26 Left Join mkpp On mkpp.npwp = pph26.NPWP_KPP " & _
                "Where pph26.kode_divisi = '" & kdDivisi & "' And " & _
                "pph26.Masa_Pajak = '" & masa2 & "' And " & _
                "pph26.Tahun_Pajak = '" & tahun & "' and PPh_Yang_Dipotong__2 > 0 " & _
                "Group By pph26.kode_divisi, pph26.NPWP_KPP, mkpp.nama, " & _
                "mkpp.alamat, pph26.Masa_Pajak, pph26.Tahun_Pajak"
    ElseIf jenisPajak = "pph26_tarif3" Then
        cust_Ref = cust_Ref & "I"
        sql = "Select pph26.NPWP_KPP, left(mkpp.nama,30), left(mkpp.alamat,50), " & _
                "f_getkotaKpp(pph26.NPWP_KPP) as k4, " & _
                "'' as k5, '411127' as k6, '103' as k7, " & _
                "'" & masa2 & "' as k8, '" & masa2 & "' as k9, " & _
                "pph26.Tahun_Pajak as k10, '' as k11, 'IDR' as k12, " & _
                "Sum(pph26.Jumlah_PPh_Yang_Dipotong) as k13, '" & cust_Ref & "' as k14, " & _
                "f_getVariabel('email') as k15, '" & ket & "' as k16, '' as k17, '' as k18, " & _
                "pph26.NPWP_KPP as k19, '' as k20, 'E' as k21 " & _
                "From pph26 Left Join mkpp On mkpp.npwp = pph26.NPWP_KPP " & _
                "Where pph26.kode_divisi = '" & kdDivisi & "' And " & _
                "pph26.Masa_Pajak = '" & masa2 & "' And " & _
                "pph26.Tahun_Pajak = '" & tahun & "' and PPh_Yang_Dipotong__3 > 0 " & _
                "Group By pph26.kode_divisi, pph26.NPWP_KPP, mkpp.nama, " & _
                "mkpp.alamat, pph26.Masa_Pajak, pph26.Tahun_Pajak"
    ElseIf jenisPajak = "pph26_tarif45" Then
        cust_Ref = cust_Ref & "J"
        sql = "Select pph26.NPWP_KPP, left(mkpp.nama,30), left(mkpp.alamat,50), " & _
                "f_getkotaKpp(pph26.NPWP_KPP) as k4, " & _
                "'' as k5, '411127' as k6, '104' as k7, " & _
                "'" & masa2 & "' as k8, '" & masa2 & "' as k9, " & _
                "pph26.Tahun_Pajak as k10, '' as k11, 'IDR' as k12, " & _
                "Sum(pph26.Jumlah_PPh_Yang_Dipotong) as k13, '" & cust_Ref & "' as k14, " & _
                "f_getVariabel('email') as k15, '" & ket & "' as k16, '' as k17, '' as k18, " & _
                "pph26.NPWP_KPP as k19, '' as k20, 'E' as k21 " & _
                "From pph26 Left Join mkpp On mkpp.npwp = pph26.NPWP_KPP " & _
                "Where pph26.kode_divisi = '" & kdDivisi & "' And " & _
                "pph26.Masa_Pajak = '" & masa2 & "' And " & _
                "pph26.Tahun_Pajak = '" & tahun & _
                "' and (PPh_Yang_Dipotong__4 + PPh_Yang_Dipotong__5) > 0 " & _
                "Group By pph26.kode_divisi, pph26.NPWP_KPP, mkpp.nama, " & _
                "mkpp.alamat, pph26.Masa_Pajak, pph26.Tahun_Pajak"
    ElseIf jenisPajak = "pph42_sewa" Then
        cust_Ref = cust_Ref & "K"
        sql = "Select " & jenisPajak & ".NPWP_KPP, left(mkpp.nama,30), left(mkpp.alamat,50), " & _
                "f_getkotaKpp(" & jenisPajak & ".NPWP_KPP) as k4, " & _
                "'' as k5, '411128' as k6, '403' as k7, " & _
                "'" & masa2 & "' as k8, '" & masa2 & "' as k9, " & _
                "" & jenisPajak & ".Tahun_Pajak as k10, '' as k11, 'IDR' as k12, " & _
                "Sum(" & jenisPajak & ".Jumlah_PPh_Yang_Dipotong) as k13, '" & cust_Ref & "' as k14, " & _
                "f_getVariabel('email') as k15, '" & ket & "' as k16, '' as k17, '' as k18, " & _
                "" & jenisPajak & ".NPWP_KPP as k19, '' as k20, 'E' as k21 " & _
                "From " & jenisPajak & " Left Join mkpp On mkpp.npwp = " & jenisPajak & ".NPWP_KPP " & _
                "Where " & jenisPajak & ".kode_divisi = '" & kdDivisi & "' And " & _
                "" & jenisPajak & ".Masa_Pajak = '" & masa2 & "' And " & _
                "" & jenisPajak & ".Tahun_Pajak = '" & tahun & "' " & _
                "Group By " & jenisPajak & ".kode_divisi, " & jenisPajak & ".NPWP_KPP, mkpp.nama, " & _
                "mkpp.alamat, " & jenisPajak & ".Masa_Pajak, " & jenisPajak & ".Tahun_Pajak"
    ElseIf jenisPajak = "pph42_konstruksi" Then
        cust_Ref = cust_Ref & "L"
        sql = "Select " & jenisPajak & ".NPWP_KPP, left(mkpp.nama,30), left(mkpp.alamat,50), " & _
                "f_getkotaKpp(" & jenisPajak & ".NPWP_KPP) as k4, " & _
                "'' as k5, '411128' as k6, '409' as k7, " & _
                "'" & masa2 & "' as k8, '" & masa2 & "' as k9, " & _
                "" & jenisPajak & ".Tahun_Pajak as k10, '' as k11, 'IDR' as k12, " & _
                "Sum(" & jenisPajak & ".Jumlah_PPh_Yang_Dipotong) as k13, '" & cust_Ref & "' as k14, " & _
                "f_getVariabel('email') as k15, '" & ket & "' as k16, '' as k17, '' as k18, " & _
                "" & jenisPajak & ".NPWP_KPP as k19, '' as k20, 'E' as k21 " & _
                "From " & jenisPajak & " Left Join mkpp On mkpp.npwp = " & jenisPajak & ".NPWP_KPP " & _
                "Where " & jenisPajak & ".kode_divisi = '" & kdDivisi & "' And " & _
                "" & jenisPajak & ".Masa_Pajak = '" & masa2 & "' And " & _
                "" & jenisPajak & ".Tahun_Pajak = '" & tahun & "' " & _
                "Group By " & jenisPajak & ".kode_divisi, " & jenisPajak & ".NPWP_KPP, mkpp.nama, " & _
                "mkpp.alamat, " & jenisPajak & ".Masa_Pajak, " & jenisPajak & ".Tahun_Pajak"
    ElseIf jenisPajak = "pph15" Then
        cust_Ref = cust_Ref & "M"
        sql = "Select " & jenisPajak & ".NPWP_KPP, left(mkpp.nama,30), left(mkpp.alamat,50), " & _
                "f_getkotaKpp(" & jenisPajak & ".NPWP_KPP) as k4, " & _
                "'' as k5, '411128' as k6, '410' as k7, " & _
                "'" & masa2 & "' as k8, '" & masa2 & "' as k9, " & _
                "" & jenisPajak & ".Tahun_Pajak as k10, '' as k11, 'IDR' as k12, " & _
                "Sum(" & jenisPajak & ".pph_dipotong) as k13, '" & cust_Ref & "' as k14, " & _
                "f_getVariabel('email') as k15, '" & ket & "' as k16, '' as k17, '' as k18, " & _
                "" & jenisPajak & ".NPWP_KPP as k19, '' as k20, 'E' as k21 " & _
                "From " & jenisPajak & " Left Join mkpp On mkpp.npwp = " & jenisPajak & ".NPWP_KPP " & _
                "Where " & jenisPajak & ".kode_divisi = '" & kdDivisi & "' And " & _
                "" & jenisPajak & ".Masa_Pajak = '" & masa2 & "' And " & _
                "" & jenisPajak & ".Tahun_Pajak = '" & tahun & "' " & _
                "Group By " & jenisPajak & ".kode_divisi, " & jenisPajak & ".NPWP_KPP, mkpp.nama, " & _
                "mkpp.alamat, " & jenisPajak & ".Masa_Pajak, " & jenisPajak & ".Tahun_Pajak"
    End If
    
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox(Err.DESCRIPTION, "sql error", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rs)
    If jRec > 0 Then
        c = 1
        rs.MoveFirst
        Do While rs.EOF = False
            Call info_progress(sb1, 2, c, jRec, "sum " & jenisPajak)
            rsData.AddNew
            For d = 0 To rsData.Fields.Count - 1
                rsData.Fields(d).Value = rs.Fields(d).Value
            Next
            rsData.Update
            rs.MoveNext
            c = c + 1
        Loop
    End If
    
    jRec = RecordCount(rsData)
    If jRec <= 0 Then
        Me.cmd_export.Enabled = False
    Else
        Me.cmd_export.Enabled = True
    End If
End Sub

Private Sub cmd_xls_Click()
    Dim jRec As Long
    
    jRec = RecordCount(rsData)
    If jRec <= 0 Then
        Call pesan2("RsData no data")
        Exit Sub
    End If
    If sudah_ada_header = False Then
        Call create_header
    End If
    
    Call create_xls3(rsGrid, "", "", "", "", "", "", False, False, False)
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
  sudah_ada_header = False
  Call set_tombol_load
End Sub

Sub set_tombol_load()
    If sudah_ada_header = False Then
        Me.cmd_export.Caption = "3. Add Header"
    ElseIf sudah_ada_header = True Then
        Me.cmd_export.Caption = "3. Export CSV"
    End If
End Sub


Private Sub Form_Resize()
    Me.Shape1.Width = Me.Width
    Me.lb_caption.Width = Me.Width
    
    If Me.Width - 645 > 0 Then Me.Frame2.Width = Me.Width - 645
    If Me.Frame2.Width - 240 > 0 Then Me.DataGrid1.Width = Me.Frame2.Width - 240
    If Me.Frame2.Width - 1695 > 0 Then Me.cmd_export.Left = Me.Frame2.Width - 1695
    
    'height
    If Me.Height - 3720 > 0 Then Me.Frame2.Height = Me.Height - 3720
    If Me.Frame2.Height - 840 > 0 Then Me.DataGrid1.Height = Me.Frame2.Height - 840
    If Me.Frame2.Height - 495 > 0 Then Me.cmd_export.Top = Me.Frame2.Height - 495
    
    Me.cmd_xls.Top = Me.cmd_export.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call dbMySQL_close
End Sub

Sub DownloadFile()
    Dim cek As String
    
    cek = tbBniDirect_get(get_kode_combo(Me.cb_divisi, "-"), "user1")
    If Trim(cek) = "" Then
        Call pesan2("user name invalid")
    Else
        Call DownloadFile2(tbVariabel_get("ftpip"), _
                            cek, _
                            tbBniDirect_get(get_kode_combo(Me.cb_divisi, "-"), "pass1"), _
                            tbVariabel_get("ftpport"), "")
    End If
    
End Sub

Sub lihat_isiFolder()
    Dim List1 As String
    Dim filesys, filepath, TargetFolder, Files, targetfile
    Dim i As Integer
    
    

    Set filesys = CreateObject("Scripting.FileSystemObject")
    filepath = filesys.GetAbsolutePathName("")

    Set TargetFolder = filesys.GetFolder(App.Path & "\exp")
    Set Files = TargetFolder.Files
    i = 1
    List1 = ""
    For Each targetfile In Files
        'objSheet.Cells(i, j).Value = targetfile.Name
        List1 = List1 & targetfile.Name & ", "
        i = i + 1
    Next
    
    
    MsgBox List1
End Sub

Private Sub mnImpTaxInqury_Click()
    frm_csv_BNI_imp.Show
End Sub

Private Sub mntaxinquiry_Click()
    frm_csv_BNI_rep.Show
End Sub
