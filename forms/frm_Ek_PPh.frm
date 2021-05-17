VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Ek_PPh 
   ClientHeight    =   7245
   ClientLeft      =   240
   ClientTop       =   1050
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
   Begin VB.Frame Frame1 
      Caption         =   " Filter "
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   12015
      Begin VB.ComboBox cb_tahun 
         Height          =   330
         Left            =   4800
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   360
         Width           =   1215
      End
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
      Begin VB.ComboBox cb_JenisPPh 
         Height          =   330
         Left            =   1200
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   4200
         TabIndex        =   9
         Top             =   420
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pajak"
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   420
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   12015
      Begin VB.CommandButton cmd_pelaporan 
         Caption         =   "Pelaporan SPT PPh"
         Height          =   375
         Left            =   7200
         TabIndex        =   11
         Top             =   4920
         Width           =   3615
      End
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
      Caption         =   "Ekualisasi : PPh"
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
      Begin VB.Menu mnRekLapor 
         Caption         =   "Rekap"
      End
   End
End
Attribute VB_Name = "frm_Ek_PPh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim nama_data As String
Dim cnnTemp As ADODB.connection


Sub disable_Form()
    Me.Frame3.Enabled = False
    Me.Frame1.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame3.Enabled = True
    Me.Frame1.Enabled = True
End Sub


Function generate_sql() As String
    Dim sql As String, kondisi As String
    Dim cari As String, jmlData As Integer
    
    'kondisi
    kondisi = ""
    cari = ""
        
    sql = "select no1 as NO_, tahun, jenis, " & _
        "kode_akun as Akun, deskripsi_akun as Deskripsi, nilai " & _
        "from ek_pph "
    
    '-- gabungkan kondisi
    If Trim(kondisi) <> "" Then
        sql = sql & " where (" & kondisi & ") "
    End If
    
    '-- gabungkan cari
    If Trim(cari) <> "" Then
        If Trim(kondisi) <> "" Then
            sql = sql & " and (" & cari & ") "
        Else
            sql = sql & " where " & cari
        End If
    End If
        
    sql = sql & " order by id1"
        
    
    'jmlData = cek_Int(InputBox("Jumlah data yang ditampilkan? (0:semua)", "", "0"))
    
    generate_sql = sql
End Function

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
            
            'kecil
            If c = 2 Or c = 3 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                If c = 3 Then
                    Me.DataGrid1.Columns(c).Width = 1200
                Else
                    Me.DataGrid1.Columns(c).Width = 800
                End If
            End If
            
            'kecil
            If c = 4 Then
                Me.DataGrid1.Columns(c).Alignment = dbgLeft
                Me.DataGrid1.Columns(c).Width = 2100
            End If
            
            'If c = 12 Or c = 20 Then
            '    Me.DataGrid1.Columns(c).Alignment = dbgCenter
            '    Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
            '    Me.DataGrid1.Columns(c).Width = 900
            'End If
    
            If c = 5 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 2500
            End If
        Next
End Sub



Private Sub cb_JenisPPh_Click()
    nama_data = Me.cb_JenisPPh.text
    Me.cmd_pelaporan.Caption = "Pelaporan SPT " & Me.cb_JenisPPh.text & " Tahun " & Me.cb_tahun.text
    Me.mnRekLapor.Caption = "Rekap Ekualisasi PPh Tahun " & Me.cb_tahun.text
End Sub

Private Sub cb_tahun_Click()
    nama_data = Me.cb_JenisPPh.text
    Me.cmd_pelaporan.Caption = "Pelaporan SPT " & Me.cb_JenisPPh.text & " Tahun " & Me.cb_tahun.text
    Me.mnRekLapor.Caption = "Rekap Ekualisasi PPh Tahun " & Me.cb_tahun.text
End Sub

Private Sub cmd_export_Click()
    Dim jRec As Long
    
    Me.disable_Form
    jRec = RecordCount(rs)
    If jRec > 0 Then
        Call create_xls2(rs, "", "005, 012, 013, 014, 015, 016, 017, 018, 019, 020, 021, 022, " & _
                        "023, 024, 025, 026, " & _
                        "027, 030, 033, 036, 039, 042, 045, 048, 051, 054, 057, 060, 063, 066, " & _
                        "069, 072, 073, 074, " & _
                        "075, 076, 077, 078, 079, 080, 081, 082, 083, 087, 090, 093, 096, 099, " & _
                        "102, 105, 108, " & _
                        "111, 114, 117, 120, 123, 126, 129, 130, 131,", "", , , , , 3)
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
        If OpenRecordSet(cnnTemp, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
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

Function fetch_data_sql_biaya(jenisPPh As String, pTahun As String) As String
    Dim sql As String
    
    If Trim(jenisPPh) = "PPh21" Then
        sql = "select tahun, kode_akun, F_get_nama_akun(kode_akun) as acct_name, jumlah " & _
            "From " & _
            "( " & _
            "select tahun, kode_akun, sum(debit) - sum(kredit)as jumlah " & _
            "From all2016_tb " & _
            "where tahun = '" & pTahun & "' " & _
            "and kode_akun in " & _
            "('51101','51111','51113','51114','51115','51116','51117','51119','51121', " & _
            "'51122','51125','51201','51206','51213','51214','51215','51216','51219', " & _
            "'51221','51222','51225','51228','51501','51502','51861','80101','80111', " & _
            "'80113','80114','80115','80116','80118','80119','80121','80123','80124', " & _
            "'80125','80126','80127','80128','80131','80133','80135','80136','80137', " & _
            "'80138','80139','80201','80203','80213','80214','80215','80216','80219', " & _
            "'80221','80223','80225','80501','80601','83261') " & _
            "group by tahun, kode_akun " & _
            ")as t "
    ElseIf Trim(jenisPPh) = "PPh22" Then
        sql = "select tahun, kode_akun, F_get_nama_akun(kode_akun) as acct_name, jumlah " & _
            "From " & _
            "( " & _
            "select tahun, kode_akun, sum(debit) - sum(kredit)as jumlah " & _
            "From all2016_tb " & _
            "where tahun = '" & pTahun & "' " & _
            "and kode_akun in " & _
            "('50201','50202') " & _
            "group by tahun, kode_akun " & _
            ")as t "
    ElseIf Trim(jenisPPh) = "PPh23" Then
        sql = "select tahun, kode_akun, F_get_nama_akun(kode_akun) as acct_name, jumlah " & _
            "From " & _
            "( " & _
            "select tahun, kode_akun, sum(debit) - sum(kredit)as jumlah " & _
            "From all2016_tb " & _
            "where tahun = '" & pTahun & "' " & _
            "and kode_akun in " & _
            "('50103','50401','50402','50405','50406','50411','50412','50413' " & _
            ",'50414','50431','50432','51401','51402','51404','51801','51802', " & _
            "'51803','51852','51853','51901','51903','71101','71102','71103', " & _
            "'81104','81121','81122','83107','83211','83212','83252','83253') " & _
            "group by tahun, kode_akun " & _
            ")as t "
    ElseIf Trim(jenisPPh) = "PPh4(2)subkon" Then
        sql = "select tahun, kode_akun, F_get_nama_akun(kode_akun) as acct_name, jumlah " & _
            "From " & _
            "( " & _
            "select tahun, kode_akun, sum(debit) - sum(kredit)as jumlah " & _
            "From all2016_tb " & _
            "where tahun = '" & pTahun & "' " & _
            "and kode_akun in " & _
            "('50301','50302','50101') " & _
            "group by tahun, kode_akun " & _
            ")as t "
    ElseIf Trim(jenisPPh) = "PPh4(2)sewa" Then
        sql = "select tahun, kode_akun, F_get_nama_akun(kode_akun) as acct_name, jumlah " & _
            "From " & _
            "( " & _
            "select tahun, kode_akun, sum(debit) - sum(kredit)as jumlah " & _
            "From all2016_tb " & _
            "where tahun = '" & pTahun & "' " & _
            "and kode_akun in " & _
            "('51851','51854','83251') " & _
            "group by tahun, kode_akun " & _
            ")as t "
    Else
        sql = "select tahun, kode_akun, F_get_nama_akun(kode_akun) as acct_name, jumlah " & _
            "From " & _
            "( " & _
            "select tahun, kode_akun, sum(debit) - sum(kredit)as jumlah " & _
            "From all2016_tb " & _
            "where tahun = '" & pTahun & "' " & _
            "and kode_akun in " & _
            "('a'') " & _
            "group by tahun, kode_akun " & _
            ")as t "
    End If
    
    fetch_data_sql_biaya = sql
End Function

Function fetch_data_sql_hutangAwal(jenisPPh As String, pTahun1 As String, _
                                    Optional isHutangAwal As Boolean = True) As String
    Dim sql As String
    Dim pTahun As String, maxBulan As String
    
    If isHutangAwal = True Then
        'tahun dikurangi 1
        pTahun = CStr(CInt(pTahun1) - 1)
    Else
        'berarti hutang akhir
        pTahun = pTahun1
    End If
    
    '-- ambil data tb utk bulan paling akhir
    sql = "select max(bulan) From all2016_tb where tahun = '" & pTahun & "'"
    maxBulan = cari_data1(cnn, sql)
    
    If Trim(jenisPPh) = "PPh21" Then
        sql = "select tahun, kode_akun, F_get_nama_akun(kode_akun) as acct_name, jumlah " & _
            "From " & _
            "( " & _
            "select tahun, kode_akun, sum(debit) - sum(kredit)as jumlah " & _
            "From all2016_tb " & _
            "where tahun = '" & pTahun & "' and bulan = '" & maxBulan & "' " & _
            "and kode_akun in " & _
            "('20701','21902','20704','20709') " & _
            "group by tahun, kode_akun " & _
            ")as t "
    ElseIf Trim(jenisPPh) = "PPh22" Then
        sql = "select tahun, kode_akun, F_get_nama_akun(kode_akun) as acct_name, jumlah " & _
            "From " & _
            "( " & _
            "select tahun, kode_akun, sum(debit) - sum(kredit)as jumlah " & _
            "From all2016_tb " & _
            "where tahun = '" & pTahun & "' and bulan = '" & maxBulan & "' " & _
            "and kode_akun in " & _
            "('20101','20102','2010200001','2010200002') " & _
            "group by tahun, kode_akun " & _
            ")as t "
    ElseIf Trim(jenisPPh) = "PPh23" Then
        sql = "select tahun, kode_akun, F_get_nama_akun(kode_akun) as acct_name, jumlah " & _
            "From " & _
            "( " & _
            "select tahun, kode_akun, sum(debit) - sum(kredit)as jumlah " & _
            "From all2016_tb " & _
            "where tahun = '" & pTahun & "' and bulan = '" & maxBulan & "' " & _
            "and kode_akun in " & _
            "('20133','20138') " & _
            "group by tahun, kode_akun " & _
            ")as t "
    ElseIf Trim(jenisPPh) = "PPh4(2)subkon" Then
        sql = "select tahun, kode_akun, F_get_nama_akun(kode_akun) as acct_name, jumlah " & _
            "From " & _
            "( " & _
            "select tahun, kode_akun, sum(debit) - sum(kredit)as jumlah " & _
            "From all2016_tb " & _
            "where tahun = '" & pTahun & "' and bulan = '" & maxBulan & "' " & _
            "and kode_akun in " & _
            "('20111','20112','20113','20131','20116') " & _
            "group by tahun, kode_akun " & _
            ")as t "
    ElseIf Trim(jenisPPh) = "PPh4(2)sewa" Then
        sql = "select tahun, kode_akun, F_get_nama_akun(kode_akun) as acct_name, jumlah " & _
            "From " & _
            "( " & _
            "select tahun, kode_akun, sum(debit) - sum(kredit)as jumlah " & _
            "From all2016_tb " & _
            "where tahun = '" & pTahun & "' and bulan = '" & maxBulan & "' " & _
            "and kode_akun in " & _
            "('') " & _
            "group by tahun, kode_akun " & _
            ")as t "
    Else
        sql = "select tahun, kode_akun, F_get_nama_akun(kode_akun) as acct_name, jumlah " & _
            "From " & _
            "( " & _
            "select tahun, kode_akun, sum(debit) - sum(kredit)as jumlah " & _
            "From all2016_tb " & _
            "where tahun = '" & pTahun & "' and bulan = '" & maxBulan & "' " & _
            "and kode_akun in " & _
            "('a'') " & _
            "group by tahun, kode_akun " & _
            ")as t "
    End If
    
    fetch_data_sql_hutangAwal = sql
End Function

Function fetch_data_jmlDPPBerdasarkanSPT(jenisPPh As String, pTahun As String) As Currency
    Dim sql As String, t As String
    
    If Trim(jenisPPh) = "PPh21" Then
        sql = "select sum(a) " & _
            "From " & _
            "( " & _
            "select sum(jumlah_bruto) as a " & _
            "From pph21bulanan " & _
            "where Tahun_pajak = '" & pTahun & "' " & _
            "Union All " & _
            "select sum(Jumlah_DPP) as a " & _
            "From pph21tf " & _
            "where Tahun_pajak = '" & pTahun & "') as t"
    ElseIf Trim(jenisPPh) = "PPh22" Then
        sql = "select sum(Nilai_DPP) " & _
            "From pph22 " & _
            "where Tahun_pajak = '" & pTahun & "'"
    ElseIf Trim(jenisPPh) = "PPh23" Then
        sql = "select sum(Jumlah_Nilai_Bruto_) " & _
            "From pph23 " & _
            "where Tahun_pajak = '" & pTahun & "' "

    ElseIf Trim(jenisPPh) = "PPh4(2)subkon" Then
        sql = "select sum(Jumlah_Nilai_Bruto_1 + Jumlah_Nilai_Bruto_2 + Jumlah_Nilai_Bruto_3 " & _
            "+ Jumlah_Nilai_Bruto_4 + Jumlah_Nilai_Bruto_5 + Jumlah_Nilai_Bruto_6 + " & _
            "Jumlah_Nilai_Bruto_7 + Jumlah_Nilai_Bruto_8) " & _
            "From pph42_konstruksi " & _
            "where Tahun_pajak = '" & pTahun & "'"

    ElseIf Trim(jenisPPh) = "PPh4(2)sewa" Then
        sql = "select sum(Jumlah_Nilai_Bruto_1 + Jumlah_Nilai_Bruto_2 + Jumlah_Nilai_Bruto_3 " & _
            "+ Jumlah_Nilai_Bruto_4 + Jumlah_Nilai_Bruto_5 + Jumlah_Nilai_Bruto_6 + " & _
            "Jumlah_Nilai_Bruto_7 + Jumlah_Nilai_Bruto_8) " & _
            "From pph42_sewa " & _
            "where Tahun_pajak = '" & pTahun & "'"
    Else
        sql = "select 0"
    End If
    
    t = cari_data1(cnn, sql)
    fetch_data_jmlDPPBerdasarkanSPT = cek_Money(t)
End Function

Function fetch_data_getTarif(jenisPPh As String) As Double
    Dim Tarif As Double
    
    '-- dalam prosen
    
    If Trim(jenisPPh) = "PPh21" Then
        Tarif = 5
    ElseIf Trim(jenisPPh) = "PPh22" Then
        Tarif = 1.5
    ElseIf Trim(jenisPPh) = "PPh23" Then
        Tarif = 2
    ElseIf Trim(jenisPPh) = "PPh4(2)subkon" Then
        Tarif = 3
    ElseIf Trim(jenisPPh) = "PPh4(2)sewa" Then
        Tarif = 10
    Else
        Tarif = 0
    End If
    fetch_data_getTarif = Tarif
End Function


Sub fetch_data(jenisPPh As String, pTahun As String)
    
    Dim sql As String, kondisi As String, t As String
    Dim rsSumber As ADODB.Recordset, rsTujuan As ADODB.Recordset
    Dim jRec As Long, c As Long, a As Integer
    
    Dim klm()
    Dim isi(), p
    
    Dim kode_akun As String, acct_name As String, Jumlah As Currency
    Dim subTotalBiaya As Currency, subTotalHutangAwal As Currency, subTotalHutangAkhir As Currency
    Dim jmlDppBerdasarkanTB As Currency, jmlDppBerdasarkanSPT As Currency, selisihDPP As Currency
    Dim Tarif As Double, kurangSetor As Currency
    
    'open dbtemp
    'cleardata
    
    
    'open cnnTemp
    Call db_Access_open(cnnTemp, App.Path & "\data\dbrep.mdb")
    klm = Array("no1", "tahun", "jenis", "kode_akun", "deskripsi_akun", "nilai")
    
    'delete
    sql = "delete from ek_pph"
    If ExecSQL1(cnnTemp, sql) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    'insert header
    If Trim(jenisPPh) = "PPh21" Then
        isi = Array("", "", "", "", "PPh Pasal 21", "0")
    ElseIf Trim(jenisPPh) = "PPh22" Then
        isi = Array("", "", "", "", "PPh Pasal 22", "0")
    ElseIf Trim(jenisPPh) = "PPh23" Then
        isi = Array("", "", "", "", "PPh Pasal 23", "0")
    ElseIf Trim(jenisPPh) = "PPh4(2)subkon" Then
        isi = Array("", "", "", "", "PPh Pasal 4(2) Subkon", "0")
    ElseIf Trim(jenisPPh) = "PPh4(2)sewa" Then
        isi = Array("", "", "", "", "PPh Pasal 4(2) Sewa", "0")
    Else
        isi = Array("", "", "", "", "error", "0")
    End If
    
    If tbInsert("ek_pph", klm, isi, cnnTemp) = False Then
        p = MsgBox("Lanjut?", vbYesNo)
        If p = vbNo Then Exit Sub
    End If
    
    'biaya
    '-- load dari tb, sesuai tahun
    sql = fetch_data_sql_biaya(jenisPPh, pTahun)
    If OpenRecordSet(cnn, rsSumber, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rsSumber)
    subTotalBiaya = 0
    If jRec > 0 Then
        rsSumber.MoveFirst
        c = 1
        Do While rsSumber.EOF = False
            Call info_progress(Me.StatusBar1, 1, c, jRec, "Load Biaya")
            kode_akun = cek_null(rsSumber(1))
            acct_name = cek_null(rsSumber(2))
            Jumlah = cek_null(rsSumber(3))
            
            isi = Array(c, pTahun, "Biaya", kode_akun, acct_name, Jumlah)
        
            If tbInsert("ek_pph", klm, isi, cnnTemp) = False Then
                p = MsgBox("Lanjut?", vbYesNo)
                If p = vbNo Then Exit Sub
            End If
            
            rsSumber.MoveNext
            c = c + 1
            subTotalBiaya = subTotalBiaya + cek_Money(Jumlah)
        Loop
    End If
    
    isi = Array("", "", "", "", "Sub Total Biaya", subTotalBiaya)
    If tbInsert("ek_pph", klm, isi, cnnTemp) = False Then
        p = MsgBox("Lanjut?", vbYesNo)
        If p = vbNo Then Exit Sub
    End If
    
    'hutang awal
    '-- load dari tb, sesuai tahun
    
    sql = fetch_data_sql_hutangAwal(jenisPPh, pTahun)
    If OpenRecordSet(cnn, rsSumber, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rsSumber)
    subTotalHutangAwal = 0
    If jRec > 0 Then
        rsSumber.MoveFirst
        Do While rsSumber.EOF = False
            Call info_progress(Me.StatusBar1, 1, c, jRec, "Load Biaya")
            kode_akun = cek_null(rsSumber(1))
            acct_name = cek_null(rsSumber(2))
            Jumlah = cek_null(rsSumber(3))
            
            isi = Array(c, CStr(CInt(pTahun) - 1), "Hutang Awal", kode_akun, acct_name, Jumlah)
        
            If tbInsert("ek_pph", klm, isi, cnnTemp) = False Then
                p = MsgBox("Lanjut?", vbYesNo)
                If p = vbNo Then Exit Sub
            End If
            
            rsSumber.MoveNext
            c = c + 1
            subTotalHutangAwal = subTotalHutangAwal + cek_Money(Jumlah)
        Loop
    End If
    isi = Array("", "", "", "", "Sub Total Hutang Awal", subTotalHutangAwal)
    If tbInsert("ek_pph", klm, isi, cnnTemp) = False Then
        p = MsgBox("Lanjut?", vbYesNo)
        If p = vbNo Then Exit Sub
    End If
    
    
    'hutang akhir
    '-- load dari tb, sesuai tahun
    sql = fetch_data_sql_hutangAwal(jenisPPh, pTahun, False)
    If OpenRecordSet(cnn, rsSumber, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rsSumber)
    subTotalHutangAkhir = 0
    If jRec > 0 Then
        rsSumber.MoveFirst
        Do While rsSumber.EOF = False
            Call info_progress(Me.StatusBar1, 1, c, jRec, "Load Biaya")
            kode_akun = cek_null(rsSumber(1))
            acct_name = cek_null(rsSumber(2))
            Jumlah = cek_null(rsSumber(3))
            
            isi = Array(c, pTahun, "Hutang Akhir", kode_akun, acct_name, Jumlah)
        
            If tbInsert("ek_pph", klm, isi, cnnTemp) = False Then
                p = MsgBox("Lanjut?", vbYesNo)
                If p = vbNo Then Exit Sub
            End If
            
            rsSumber.MoveNext
            c = c + 1
            subTotalHutangAkhir = subTotalHutangAkhir + cek_Money(Jumlah)
        Loop
    End If
    isi = Array("", "", "", "", "Sub Total Hutang Akhir", subTotalHutangAkhir)
    If tbInsert("ek_pph", klm, isi, cnnTemp) = False Then
        p = MsgBox("Lanjut?", vbYesNo)
        If p = vbNo Then Exit Sub
    End If
    
    jmlDppBerdasarkanTB = subTotalBiaya + subTotalHutangAwal - subTotalHutangAkhir
    isi = Array("", "", "", "", "Jumlah DPP Berdasarkan TB", jmlDppBerdasarkanTB)
    If tbInsert("ek_pph", klm, isi, cnnTemp) = False Then
        p = MsgBox("Lanjut?", vbYesNo)
        If p = vbNo Then Exit Sub
    End If
    
    jmlDppBerdasarkanSPT = fetch_data_jmlDPPBerdasarkanSPT(jenisPPh, pTahun)
    isi = Array("", "", "", "", "Jumlah DPP Berdasarkan SPT", jmlDppBerdasarkanSPT)
    If tbInsert("ek_pph", klm, isi, cnnTemp) = False Then
        p = MsgBox("Lanjut?", vbYesNo)
        If p = vbNo Then Exit Sub
    End If
    
    selisihDPP = jmlDppBerdasarkanTB - jmlDppBerdasarkanSPT
    isi = Array("", "", "", "", "Selisih DPP ", selisihDPP)
    If tbInsert("ek_pph", klm, isi, cnnTemp) = False Then
        p = MsgBox("Lanjut?", vbYesNo)
        If p = vbNo Then Exit Sub
    End If
    
    Tarif = fetch_data_getTarif(jenisPPh)
    isi = Array("", "", "", "", "Tarif: " & CStr(Tarif) & "%", "0")
    If tbInsert("ek_pph", klm, isi, cnnTemp) = False Then
        p = MsgBox("Lanjut?", vbYesNo)
        If p = vbNo Then Exit Sub
    End If
    
    If selisihDPP > 0 Then
        kurangSetor = selisihDPP * Tarif / 100
    Else
        kurangSetor = 0
    End If
    isi = Array("", "", "", "", "Kurang Setor", kurangSetor)
    If tbInsert("ek_pph", klm, isi, cnnTemp) = False Then
        p = MsgBox("Lanjut?", vbYesNo)
        If p = vbNo Then Exit Sub
    End If
    
End Sub

Function fetch_pelaporan_pph_sql(jenisPPh As String, pTahun As String)
    Dim sql As String
    
    If Trim(jenisPPh) = "PPh21" Then
        sql = "select NPWP_KPP as NPWP,F_get_nama_kpp(NPWP_KPP) as KPP, sum(dpp), sum(pph) " & _
            "From " & _
            "( " & _
            "select NPWP_KPP, sum(jumlah_bruto) as dpp, sum(jumlah_pph) as pph " & _
            "From pph21bulanan " & _
            "where Tahun_pajak = '" & pTahun & "' " & _
            "group by NPWP_KPP " & _
            "Union All " & _
            "select NPWP_KPP, sum(jumlah_dpp) as dpp, sum(jumlah_pph) as pph " & _
            "From pph21tf " & _
            "where Tahun_pajak = '" & pTahun & "' " & _
            "group by NPWP_KPP " & _
            ") as t " & _
            "group by NPWP_KPP " & _
            "Union All " & _
            "select '' as NPWP_KPP, 'TOTAL' as KPP, sum(dpp) as dpp, sum(pph) as pph " & _
            "From " & _
            "( " & _
            "select sum(jumlah_bruto) as dpp, sum(Jumlah_PPh) as pph " & _
            "From pph21bulanan " & _
            "where Tahun_pajak = '" & pTahun & "' " & _
            "Union All " & _
            "select sum(Jumlah_DPP) as dpp, sum(Jumlah_PPh) as pph " & _
            "From pph21tf "
        sql = sql & _
            "where Tahun_pajak = '" & pTahun & "') as t"
    ElseIf Trim(jenisPPh) = "PPh22" Then
        sql = "select NPWP_KPP, F_get_nama_kpp(NPWP_KPP) as KPP, sum(Nilai_DPP) as dpp,  " & _
            "sum(Nilai_PPh)  as pph " & _
            "From pph22 " & _
            "where Tahun_pajak = '" & pTahun & "' " & _
            "group by NPWP_KPP " & _
            "Union All " & _
            "select '' as NPWP_KPP, 'TOTAL' as KPP, sum(Nilai_DPP) as dpp, sum(Nilai_PPh) as pph " & _
            "From pph22 " & _
            "where Tahun_pajak = '" & pTahun & "'"
            
    ElseIf Trim(jenisPPh) = "PPh23" Then
        sql = "select NPWP_KPP, F_get_nama_kpp(NPWP_KPP) as KPP, " & _
            "sum(Jumlah_Nilai_Bruto_) as dpp,  sum(Jumlah_PPh_Yang_Dipotong)  as pph " & _
            "From pph23 " & _
            "where Tahun_pajak = '" & pTahun & "' " & _
            "group by NPWP_KPP " & _
            "Union All " & _
            "select '' as NPWP_KPP, 'TOTAL' as KPP, sum(Jumlah_Nilai_Bruto_) as dpp, " & _
            "sum(Jumlah_PPh_Yang_Dipotong) As PPh " & _
            "From pph23 " & _
            "where Tahun_pajak = '" & pTahun & "'"

    ElseIf Trim(jenisPPh) = "PPh4(2)subkon" Then
        sql = "select NPWP_KPP, F_get_nama_kpp(NPWP_KPP) as KPP, " & _
            "sum(Jumlah_Nilai_Bruto_1 + Jumlah_Nilai_Bruto_2 + Jumlah_Nilai_Bruto_3 " & _
            "+ Jumlah_Nilai_Bruto_4 + Jumlah_Nilai_Bruto_5 + Jumlah_Nilai_Bruto_6 + " & _
            "Jumlah_Nilai_Bruto_7 + Jumlah_Nilai_Bruto_8) as dpp,  sum(Jumlah_PPh_Yang_Dipotong)  " & _
            "as pph From pph42_konstruksi " & _
            "where Tahun_pajak = '" & pTahun & "' " & _
            "group by NPWP_KPP " & _
            "Union All " & _
            "select '' as NPWP_KPP, 'TOTAL' as KPP, sum(Jumlah_Nilai_Bruto_1 +  " & _
            "Jumlah_Nilai_Bruto_2 + Jumlah_Nilai_Bruto_3 " & _
            "+ Jumlah_Nilai_Bruto_4 + Jumlah_Nilai_Bruto_5 + Jumlah_Nilai_Bruto_6 + " & _
            "Jumlah_Nilai_Bruto_7 + Jumlah_Nilai_Bruto_8) as dpp, " & _
            "sum(Jumlah_PPh_Yang_Dipotong) As PPh " & _
            "From pph42_konstruksi " & _
            "where Tahun_pajak = '" & pTahun & "'"

    ElseIf Trim(jenisPPh) = "PPh4(2)sewa" Then
        sql = "select NPWP_KPP, F_get_nama_kpp(NPWP_KPP) as KPP, " & _
            "sum(Jumlah_Nilai_Bruto_1 + Jumlah_Nilai_Bruto_2 + Jumlah_Nilai_Bruto_3 " & _
            "+ Jumlah_Nilai_Bruto_4 + Jumlah_Nilai_Bruto_5 + Jumlah_Nilai_Bruto_6 + " & _
            "Jumlah_Nilai_Bruto_7 + Jumlah_Nilai_Bruto_8) as dpp,  sum(Jumlah_PPh_Yang_Dipotong)  " & _
            "as pph From pph42_sewa " & _
            "where Tahun_pajak = '" & pTahun & "' " & _
            "group by NPWP_KPP " & _
            "Union All " & _
            "select '' as NPWP_KPP, 'TOTAL' as KPP, sum(Jumlah_Nilai_Bruto_1 +  " & _
            "Jumlah_Nilai_Bruto_2 + Jumlah_Nilai_Bruto_3 " & _
            "+ Jumlah_Nilai_Bruto_4 + Jumlah_Nilai_Bruto_5 + Jumlah_Nilai_Bruto_6 + " & _
            "Jumlah_Nilai_Bruto_7 + Jumlah_Nilai_Bruto_8) as dpp, " & _
            "sum(Jumlah_PPh_Yang_Dipotong) As PPh " & _
            "From pph42_sewa " & _
            "where Tahun_pajak = '" & pTahun & "'"
    Else
        sql = "select 0"
    End If
    fetch_pelaporan_pph_sql = sql
End Function


Private Sub cmd_load_Click()
    Call fetch_data(Me.cb_JenisPPh.text, Me.cb_tahun.text)
    Call LoadGrid
End Sub


Private Sub cmd_pelaporan_Click()
    Dim sql As String
    
    frm_Grid.Show
    sql = fetch_pelaporan_pph_sql(Me.cb_JenisPPh.text, Me.cb_tahun.text)
    frm_Grid.sql = sql
    frm_Grid.judul = "Rekap Pelaporan SPT " & Me.cb_JenisPPh.text & " Tahun " & Me.cb_tahun.text
    Call frm_Grid.LoadGrid(2, 3)
End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
  Call dbMySQL_open
    
  'load combo
  Me.cb_JenisPPh.Clear
  Me.cb_JenisPPh.AddItem "PPh21"
  Me.cb_JenisPPh.AddItem "PPh22"
  Me.cb_JenisPPh.AddItem "PPh23"
  Me.cb_JenisPPh.AddItem "PPh4(2)subkon"
  Me.cb_JenisPPh.AddItem "PPh4(2)sewa"
  
  sql = "select distinct tahun from all2016_tb"
  Call Load_combo(Me.cb_tahun, sql, cnn, True)
  
  
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
 
 'Call LoadGrid
  Call pesan2("klik cari data dan ENTER")
End Sub


Private Sub Form_Resize()
    If Me.Width - 405 > 0 Then Me.Frame3.Width = Me.Width - 405
    If Me.Height - 2595 > 0 Then Me.Frame3.Height = Me.Height - 2595

    If Me.Width - 645 > 0 Then Me.DataGrid1.Width = Me.Width - 645
    If Me.Height - 3435 > 0 Then Me.DataGrid1.Height = Me.Height - 3435

    If Me.Height - 3090 > 0 Then Me.cmd_export.Top = Me.Height - 3090
    Me.cmd_pelaporan.Top = Me.cmd_export.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call dbMySQL_close
End Sub

Private Sub mnRekLapor_Click()
    Dim sql As String, t As String
    Dim rs As ADODB.Recordset
    Dim c As Integer
    
    '-- test
    'sql = "call P_ekPPhRekap21('" & Me.cb_tahun.text & "')"
    'If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
    '    Exit Sub
    'End If
    
    'rs.MoveFirst
    'Me.List1.Clear
    'Do While rs.EOF = False
    '    t = ""
    '    For c = 0 To rs.Fields.Count - 1
    '        t = t & rs.Fields(c).Value & ", "
    '    Next
    '    Me.List1.AddItem t
    '    rs.MoveNext
    'Loop
    '-
    
    'Exit Sub
    
    frm_Grid.Show
    sql = "call P_ekPPhRekap21('" & Me.cb_tahun.text & "')"
    frm_Grid.sql = sql
    frm_Grid.judul = "Rekap Ekualisasi PPh Tahun " & Me.cb_tahun.text
    Call frm_Grid.LoadGrid(1, 8)
End Sub
