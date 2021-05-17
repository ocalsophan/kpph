VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Ek_AllSin 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4935
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
   ScaleHeight     =   4935
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_stopLoad 
      BackColor       =   &H00C0FFC0&
      Caption         =   "stop_Load"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   " Data yang akan di sinkronisasi / di Rekap "
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   12015
      Begin VB.ComboBox cb_proyek 
         Height          =   330
         Left            =   1920
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox ch_sinWIP 
         Caption         =   "Data WIP dsb"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   2175
      End
      Begin VB.CheckBox ch_sinBP 
         Caption         =   "Data BP"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CheckBox ch_sinFP 
         Caption         =   "Data FP"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CheckBox ch_sin_master 
         Caption         =   "Data Master + PU"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton cmd_load 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Load"
         Height          =   375
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Untuk Jenis Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Untuk Proyek (lama) "
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   540
         Width           =   1485
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4680
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
      Caption         =   "Ekualisasi : Sinkronisasi"
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
Attribute VB_Name = "frm_Ek_AllSin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim stopLoad As Boolean

Sub disable_Form()
    'Me.Frame3.Enabled = False
    Me.Frame1.Enabled = False
End Sub

Sub Enable_Form()
    'Me.Frame3.Enabled = True
    Me.Frame1.Enabled = True
End Sub

Sub updateStopLoad()
    If stopLoad = True Then
        Me.cmd_stopLoad.Visible = False
    Else
        Me.cmd_stopLoad.Visible = True
    End If
End Sub


Sub sin_master()
    Dim sql As String, t As String
    Dim rs As ADODB.Recordset, p
    Dim jRec As Long, c As Long, mod1 As Integer
    Dim tahunAwal As Integer, tahunAkhir As Integer, tahun As Integer
    Dim max_bulan As String, nilai_pu As Currency
    
    Dim CABANG As String, DIVISI As String, NO_KONTRAK As String
    Dim NK_PPN As Currency, OWNER As String, PROYEK As String
    Dim KODE_ACPAC As String, KODE_PROYEK_LAMA As String, KODE_PROYEK_BARU As String
    
    'looping dari master proyek...
    'setiap kode_proyek_lama, kode_akun,
    '-- insert_update di all2016_all
    'Update pu
    
    tahunAwal = 2008
    tahunAkhir = 2022
    
    If Trim(Me.cb_proyek.text) = "ALL" Then
        sql = "select CABANG, divisi, NO_KONTRAK, " & _
            "NK_PPN, OWNER, PROYEK,  " & _
            "KODE_ACPAC, kode_Proyek_lama, kode_Proyek_baru " & _
            "from all2016_master"
    Else
        sql = "select CABANG, divisi, NO_KONTRAK, " & _
            "NK_PPN, OWNER, PROYEK,  " & _
            "KODE_ACPAC, kode_Proyek_lama, kode_Proyek_baru " & _
            "from all2016_master " & _
            "where kode_proyek_lama = '" & Trim(Me.cb_proyek.text) & "'"
    End If
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
    rs.MoveFirst
    c = 1
    stopLoad = False
    updateStopLoad
    Do While rs.EOF = False
        Call info_progress(Me.StatusBar1, 1, c, jRec, "Master Proyek Ekualisasi")
        If stopLoad = True Then
            Call pesan2("Loading di hentikan")
            Exit Do
        End If
        
        mod1 = c Mod 50
        If mod1 = 0 Then Call dbMySQL_open
        
        CABANG = cek_null(rs(0))
        DIVISI = cek_null(rs(1))
        NO_KONTRAK = cek_null(rs(2))
        NK_PPN = cek_Money(cek_null(rs(3)))
        OWNER = cek_null(rs(4))
        PROYEK = cek_null(rs(5))
        KODE_ACPAC = cek_null(rs(6))
        KODE_PROYEK_LAMA = cek_null(rs(7))
        KODE_PROYEK_BARU = cek_null(rs(8))
        
        sql = "select F_sin_ek_master('" & CABANG & "','" & DIVISI & "','" & NO_KONTRAK & _
            "','" & NK_PPN & "','" & OWNER & "','" & PROYEK & "','" & KODE_ACPAC & _
            "','" & KODE_PROYEK_LAMA & "','" & KODE_PROYEK_BARU & "')"
        t = cari_data1(cnn, sql)
        
        '-- update PU
        '-- cari di TB, dengan akun kepala 40, utk tahun 2008 s/d 2022
        For tahun = tahunAwal To tahunAkhir
            Call info_progress(Me.StatusBar1, 2, CLng(tahun), CLng(tahunAkhir), "update PU")
            sql = "select max(bulan) from all2016_tb where tahun = '" & tahun & _
                    "' and kode_proyek_lama = '" & KODE_PROYEK_LAMA & "'"
            max_bulan = cari_data1(cnn, sql)
            sql = "select sum(debit) - sum(kredit) from all2016_tb where kode_proyek_lama = '" & _
                    KODE_PROYEK_LAMA & "' and tahun = '" & tahun & "' and bulan = '" & _
                    max_bulan & "' and left(kode_akun,2) = '40'"
            t = cari_data1(cnn, sql)
            nilai_pu = cek_Money(t)
            
            sql = "update all2016_all set PU_" & tahun & "='" & nilai_pu & "' where kode_proyek_lama = '" & _
                    KODE_PROYEK_LAMA & "' and AKUN = '" & KODE_ACPAC & "'"
            If ExecSQL1(cnn, sql) <> 0 Then
                sql = InputBox("error sql", "", sql)
                p = MsgBox("Stop Load", vbYesNo)
                If p = vbYes Then
                    stopLoad = True
                    Exit Sub
                Else
                    'reconnect
                    Call dbMySQL_open
                End If
            End If
        Next
        
        c = c + 1
        rs.MoveNext
    Loop
    
    Call pesan2("Selesai")
    stopLoad = True
    Call updateStopLoad
End Sub

Sub sin_FP()
    Dim sql As String, t As String
    Dim rs As ADODB.Recordset, p, rS2 As ADODB.Recordset
    Dim jRec As Long, c As Long
    
    Dim KODE_PROYEK_LAMA As String, KODE_PROYEK_BARU As String, tahun As String
    Dim tgl_fp As Date, no_fp As String, dpp As Currency
    
    Dim NO1 As String, CABANG As String, DIVISI As String, NO_KONTRAK As String
    Dim nama_proyek As String, id1 As String
    
    'looping dari all2016_fp
    
    If Trim(Me.cb_proyek.text) = "ALL" Then
    
        sql = "select kode_proyek_lama, kode_proyek_baru, tahun, tgl_fp, no_fp, dpp " & _
            "from all2016_fp where not(tahun is null or tahun = '') "
    Else
        sql = "select kode_proyek_lama, kode_proyek_baru, tahun, tgl_fp, no_fp, dpp " & _
            "from all2016_fp where not(tahun is null or tahun = '') " & _
            "and kode_proyek_lama = '" & Trim(Me.cb_proyek.text) & "'"
    End If
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
    rs.MoveFirst
    c = 1
    stopLoad = False
    updateStopLoad
    Do While rs.EOF = False
        Call info_progress(Me.StatusBar1, 1, c, jRec, "FP")
        If c Mod 1000 = 0 Then Call dbMySQL_open
        If stopLoad = True Then
            Call pesan2("Loading di hentikan")
            Exit Do
        End If
        KODE_PROYEK_LAMA = cek_null(rs(0))
        
        'If kode_proyek_lama = "781610" Then
        '    MsgBox ("a")
        'End If
        
        KODE_PROYEK_BARU = cek_null(rs(1))
        tahun = cek_null(rs(2))
        tgl_fp = cek_Date(cek_null(rs(3)))
        no_fp = cek_null(rs(4))
        dpp = cek_Money(cek_null(rs(5)))
        
        'cek dulu, noFP ini apa sudah ada di all2016_all,
        'jika belum ada,
        '-- get no1, cabang, divisi, NO_KONTRAK, NAMA_PROYEK, KODE_PROYEK_LAMA, KODE_PROYEK_BARU--- limit 1 order by id1
        '-- insert baru... no1,  cabang, divisi, NO_KONTRAK, NAMA_PROYEK, KODE_PROYEK_LAMA, KODE_PROYEK_BARU,
        '-- tgl_FPxxx , no_fpxxx, dpp_ppnxxx
        'jika sudah ada, update
        '-- update,...tgl_FPxxx, no_fpxxx, dpp_ppnxxx
        
            
        sql = "select id1 from all2016_all where KODE_PROYEK_LAMA = '" & KODE_PROYEK_LAMA & _
            "' and KODE_PROYEK_BARU = '" & KODE_PROYEK_BARU & "' and tgl_fp" & tahun & _
            " = '" & set_tgl_perv(tgl_fp) & "' and no_fp" & tahun & " = '" & no_fp & "'"
        t = cari_data1(cnn, sql)
        If Trim(t) = "" Then
            '-- data belum ada, get data pendukung
            sql = "select NO1, CABANG, DIVISI, NO_KONTRAK, nama_proyek, id1 from all2016_all " & _
                    "where kode_proyek_lama = '" & KODE_PROYEK_LAMA & _
                    "' and kode_proyek_baru = '" & KODE_PROYEK_BARU & "' order by id1 limit 1"
            If OpenRecordSet(cnn, rS2, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
                sql = InputBox("sql error", "", sql)
                Exit Sub
            End If
            If RecordCount(rS2) > 0 Then
                rS2.MoveFirst
                NO1 = cek_null(rS2(0))
                CABANG = cek_null(rS2(1))
                DIVISI = cek_null(rS2(2))
                NO_KONTRAK = cek_null(rS2(3))
                nama_proyek = cek_null(rS2(4))
                id1 = cek_null(rS2(5))
                
                'cek dulu, jika di baris awal data nofp masih kosong, maka update..
                sql = "select id1 from all2016_all where " & _
                    "kode_proyek_lama = '" & KODE_PROYEK_LAMA & _
                    "' and kode_proyek_baru = '" & KODE_PROYEK_BARU & "' and " & _
                    " (no_fp" & tahun & " Is Null Or no_fp" & tahun & " = '') order by id1 limit 1"
                id1 = cari_data1(cnn, sql)
                If Trim(id1) <> "" Then
                    sql = "update all2016_all set tgl_fp" & tahun & " = '" & set_tgl_perv(tgl_fp) & _
                        "', no_fp" & tahun & "='" & no_fp & "', dpp_ppn" & tahun & "='" & _
                        dpp & "' where id1 = '" & id1 & "'"
                Else
                    sql = "insert into all2016_all (NO1, CABANG, DIVISI, " & _
                        "NO_KONTRAK, NAMA_PROYEK, KODE_PROYEK_LAMA, " & _
                        "KODE_PROYEK_BARU, tgl_fp" & tahun & ", no_fp" & tahun & ", " & _
                        "dpp_ppn" & tahun & ") values ('" & _
                        NO1 & "','" & CABANG & "','" & DIVISI & "','" & _
                        NO_KONTRAK & "','" & nama_proyek & "','" & KODE_PROYEK_LAMA & "','" & _
                        KODE_PROYEK_BARU & "','" & set_tgl_perv(tgl_fp) & "','" & no_fp & "','" & _
                        dpp & "')"
                End If
                If ExecSQL1(cnn, sql) <> 0 Then
                    sql = InputBox("error sql", "", sql)
                    p = MsgBox("Stop Load", vbYesNo)
                    If p = vbYes Then
                        stopLoad = True
                    Else
                        Call dbMySQL_open
                    End If
                End If
            End If
        Else
            'update
            sql = "update all2016_all set tgl_fp" & tahun & " = '" & set_tgl_perv(tgl_fp) & _
                    "', no_fp" & tahun & "='" & no_fp & "', dpp_ppn" & tahun & "='" & _
                    dpp & "' where id1 = '" & t & "'"
            If ExecSQL1(cnn, sql) <> 0 Then
                sql = InputBox("error sql", "", sql)
                p = MsgBox("Stop Load", vbYesNo)
                If p = vbYes Then stopLoad = True
            End If
        End If
        
        c = c + 1
        rs.MoveNext
    Loop
    
    Call pesan2("Selesai")
    stopLoad = True
    Call updateStopLoad
End Sub

Sub sin_WIP()
    'hitung
    '-- loopimg utk semua proyek yang ada di all2016 : KODE_PROYEK_LAMA + KODE_PROYEK_BARU + AKUN
    '-- get min id1
    '---- looping dari tahun awal s/d tahun akhir..
    '------- hitung jumlah PU
    '------- hitung total_dpp_ppn
    '------- SELISIH_pu_dppFP
    
    Dim sql As String, t As String
    Dim jRec As Long, c As Long
    Dim rsProyek As ADODB.Recordset
    Dim tahun As Integer, tahun_awal As Integer, tahun_akhir As Integer
    Dim jmlPU As Currency, total_dpp_ppn As Currency, SELISIH_pu_dppFP As Currency
    Dim piutang11101 As Currency, piutang11102 As Currency, piutang11103 As Currency
    Dim TOTAL_WIP_1 As Currency, PRESTASI_YAMP As Currency, KARYA_YDF As Currency
    Dim TOTAL_WIP_2 As Currency, TOTAL_WIP As Currency, selisih_dppfp_wip As Currency
    Dim total_dpp_PPh As Currency, dpp_fp_min_dpp_pph As Currency
    
    Dim id1 As String, KODE_PROYEK_LAMA As String, KODE_PROYEK_BARU As String
    Dim p
    
    tahun_awal = 2008
    tahun_akhir = 2022
    
    If Trim(Me.cb_proyek.text) = "ALL" Then
        sql = "select KODE_PROYEK_LAMA, KODE_PROYEK_BARU, min(id1) from all2016_all " & _
            "group by KODE_PROYEK_LAMA, KODE_PROYEK_BARU "
    Else
        sql = "select KODE_PROYEK_LAMA, KODE_PROYEK_BARU, min(id1) from all2016_all " & _
            "where kode_proyek_lama = '" & Trim(Me.cb_proyek.text) & "' " & _
            "group by KODE_PROYEK_LAMA, KODE_PROYEK_BARU "
    End If
            
    
    
    If OpenRecordSet(cnn, rsProyek, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rsProyek)
    If jRec <= 0 Then
        Call pesan2("no data proyek in all2016_all")
        Exit Sub
    End If
    
    rsProyek.MoveFirst
    c = 1
    stopLoad = False
    updateStopLoad
    Do While rsProyek.EOF = False
        Call info_progress(Me.StatusBar1, 1, c, jRec, "Hitung per Proyek")
        If c Mod 1000 = 0 Then Call dbMySQL_open
        If stopLoad = True Then
            Call pesan2("Loading di hentikan")
            Exit Do
        End If
        
        KODE_PROYEK_LAMA = cek_null(rsProyek(0))
        KODE_PROYEK_BARU = cek_null(rsProyek(1))
        id1 = cek_null(rsProyek(2))
        
        If Trim(id1) <> "" Then
            sql = "select PU_2008 + PU_2009 + PU_2010 + PU_2011 + PU_2012 + PU_2013 " & _
                " + PU_2014 + PU_2015 + PU_2016 + PU_2017 + PU_2018 + PU_2019 + PU_2020 " & _
                " + PU_2021 + PU_2022 from all2016_all where id1 = '" & id1 & "'"
            t = cari_data1(cnn, sql)
            jmlPU = cek_Money(t)
            
            sql = "select dpp_ppn2008 + dpp_ppn2009 + dpp_ppn2010 + dpp_ppn2011 + dpp_ppn2012 + dpp_ppn2013 " & _
                " + dpp_ppn2014 + dpp_ppn2015 + dpp_ppn2016 + dpp_ppn2017 + dpp_ppn2018 + dpp_ppn2019 + dpp_ppn2020 " & _
                " + dpp_ppn2021 + dpp_ppn2022 from all2016_all where id1 = '" & id1 & "'"
            t = cari_data1(cnn, sql)
            total_dpp_ppn = cek_Money(t)
            
            SELISIH_pu_dppFP = jmlPU - total_dpp_ppn
            
            sql = "SELECT `F_get_piutang`('" & KODE_PROYEK_LAMA & "', '" & KODE_PROYEK_BARU & _
                    "', '11101')"
            t = cari_data1(cnn, sql)
            piutang11101 = cek_Money(t)
            
            sql = "SELECT `F_get_piutang`('" & KODE_PROYEK_LAMA & "', '" & KODE_PROYEK_BARU & _
                    "', '11102')"
            t = cari_data1(cnn, sql)
            piutang11102 = cek_Money(t)
            
            sql = "SELECT `F_get_piutang`('" & KODE_PROYEK_LAMA & "', '" & KODE_PROYEK_BARU & _
                    "', '11103')"
            t = cari_data1(cnn, sql)
            piutang11103 = cek_Money(t)
            
            TOTAL_WIP_1 = piutang11101 + piutang11102 + piutang11103
            
            '--------
            sql = "SELECT `F_get_piutang`('" & KODE_PROYEK_LAMA & "', '" & KODE_PROYEK_BARU & _
                    "', '11601')"
            t = cari_data1(cnn, sql)
            PRESTASI_YAMP = cek_Money(t)
            
            
            sql = "SELECT `F_get_piutang`('" & KODE_PROYEK_LAMA & "', '" & KODE_PROYEK_BARU & _
                    "', '21201')"
            t = cari_data1(cnn, sql)
            KARYA_YDF = cek_Money(t)
            
            TOTAL_WIP_2 = PRESTASI_YAMP - KARYA_YDF
            
            
            TOTAL_WIP = TOTAL_WIP_1 + TOTAL_WIP_2
            selisih_dppfp_wip = total_dpp_ppn - TOTAL_WIP
            
            
            sql = "select dpp_pph2008 + dpp_pph2009 + dpp_pph2010 + dpp_pph2011 + dpp_pph2012 + dpp_pph2013 " & _
                " + dpp_pph2014 + dpp_pph2015 + dpp_pph2016 + dpp_pph2017 + dpp_pph2018 + dpp_pph2019 + dpp_pph2020 " & _
                " + dpp_pph2021 + dpp_pph2022 from all2016_all where id1 = '" & id1 & "'"
            t = cari_data1(cnn, sql)
            total_dpp_PPh = cek_Money(t)
            
            dpp_fp_min_dpp_pph = total_dpp_ppn - total_dpp_PPh
            
            '-- update
            sql = "update all2016_all set Jml_PU = '" & jmlPU & "' , total_dpp_ppn = '" & total_dpp_ppn & "', " & _
                "SELISIH_pu_dppFP = '" & SELISIH_pu_dppFP & "', piutang11101 = '" & piutang11101 & "', " & _
                "piutang11102 = '" & piutang11102 & "', piutang11103 = '" & piutang11103 & "', " & _
                "TOTAL_WIP_1 = '" & TOTAL_WIP_1 & "', PRESTASI_YAMP = '" & PRESTASI_YAMP & "', " & _
                "KARYA_YDF = '" & KARYA_YDF & "', TOTAL_WIP_2 = '" & TOTAL_WIP_2 & "', " & _
                "TOTAL_WIP = '" & TOTAL_WIP & "', selisih_dppfp_wip = '" & selisih_dppfp_wip & "', " & _
                "total_dpp_PPh = '" & total_dpp_PPh & "', dpp_fp_min_dpp_pph = '" & dpp_fp_min_dpp_pph & "' " & _
                "where id1 = '" & id1 & "'"
            If ExecSQL1(cnn, sql) <> 0 Then
                sql = InputBox("error sql", "", sql)
                p = MsgBox("Stop Load", vbYesNo)
                If p = vbYes Then stopLoad = True
            End If
        End If
        
        '---
        c = c + 1
        rsProyek.MoveNext
    Loop
    Call pesan2("Selesai sin WIP")
    stopLoad = True
    Call updateStopLoad
End Sub


Sub sin_BP()
    Dim sql As String, t As String
    Dim rs As ADODB.Recordset, p, rS2 As ADODB.Recordset
    Dim jRec As Long, c As Long
    
    Dim KODE_PROYEK_LAMA As String, KODE_PROYEK_BARU As String, tahun As String
    Dim tgl_BP As Date, no_BP As String, dpp As Currency
    
    Dim NO1 As String, CABANG As String, DIVISI As String, NO_KONTRAK As String
    Dim nama_proyek As String, id1 As String
    
    'looping dari all2016_fp
    
    If Trim(Me.cb_proyek.text) = "ALL" Then
        sql = "select kode_proyek_lama, kode_proyek_baru, tahun, tgl_bp, no_bp, dpp " & _
            "from all2016_bp where not(tahun is null or tahun = '')"
    Else
        sql = "select kode_proyek_lama, kode_proyek_baru, tahun, tgl_bp, no_bp, dpp " & _
            "from all2016_bp where not(tahun is null or tahun = '') " & _
            "and kode_proyek_lama = '" & Trim(Me.cb_proyek.text) & "'"
    End If
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
    rs.MoveFirst
    c = 1
    stopLoad = False
    updateStopLoad
    Do While rs.EOF = False
        Call info_progress(Me.StatusBar1, 1, c, jRec, "bukti Potong")
        If c Mod 1000 = 0 Then Call dbMySQL_open
        If stopLoad = True Then
            Call pesan2("Loading di hentikan")
            Exit Do
        End If
        KODE_PROYEK_LAMA = cek_null(rs(0))
        KODE_PROYEK_BARU = cek_null(rs(1))
        tahun = cek_null(rs(2))
        tgl_BP = cek_Date(cek_null(rs(3)))
        no_BP = cek_null(rs(4))
        dpp = cek_Money(cek_null(rs(5)))
        
        'cek dulu, noBP ini apa sudah ada di all2016_all,
        'jika belum ada,
        '-- get no1, cabang, divisi, NO_KONTRAK, NAMA_PROYEK, KODE_PROYEK_LAMA, KODE_PROYEK_BARU--- limit 1 order by id1
        '-- insert baru... no1,  cabang, divisi, NO_KONTRAK, NAMA_PROYEK, KODE_PROYEK_LAMA, KODE_PROYEK_BARU,
        '-- tgl_BPxxx , no_BPxxx, dpp_ppnxxx
        'jika sudah ada, update
        '-- update,...tgl_BPxxx, no_BPxxx, dpp_ppnxxx
        
            
        sql = "select id1 from all2016_all where KODE_PROYEK_LAMA = '" & KODE_PROYEK_LAMA & _
            "' and KODE_PROYEK_BARU = '" & KODE_PROYEK_BARU & "' and tgl_fp" & tahun & _
            " = '" & set_tgl_perv(tgl_BP) & "' and no_fp" & tahun & " = '" & no_BP & "'"
        t = cari_data1(cnn, sql)
        If Trim(t) = "" Then
            '-- data belum ada, get data pendukung
            sql = "select NO1, CABANG, DIVISI, NO_KONTRAK, nama_proyek, id1 from all2016_all " & _
                    "where kode_proyek_lama = '" & KODE_PROYEK_LAMA & _
                    "' and kode_proyek_baru = '" & KODE_PROYEK_BARU & "' order by id1 limit 1"
            If OpenRecordSet(cnn, rS2, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
                sql = InputBox("sql error", "", sql)
                Exit Sub
            End If
            If RecordCount(rS2) > 0 Then
                rS2.MoveFirst
                NO1 = cek_null(rS2(0))
                CABANG = cek_null(rS2(1))
                DIVISI = cek_null(rS2(2))
                NO_KONTRAK = cek_null(rS2(3))
                nama_proyek = cek_null(rS2(4))
                id1 = cek_null(rS2(5))
                
                'cek dulu, jika di baris awal data bp masih kosong, jika ya, maka update..
                sql = "select id1 from all2016_all where " & _
                    "kode_proyek_lama = '" & KODE_PROYEK_LAMA & _
                    "' and kode_proyek_baru = '" & KODE_PROYEK_BARU & "' and " & _
                    " (no_bp" & tahun & " Is Null Or no_bp" & tahun & " = '') order by id1 limit 1"
                id1 = cari_data1(cnn, sql)
                If Trim(id1) <> "" Then
                    sql = "update all2016_all set tgl_bp" & tahun & " = '" & set_tgl_perv(tgl_BP) & _
                        "', no_bp" & tahun & "='" & no_BP & "', dpp_pph" & tahun & "='" & _
                        dpp & "' where id1 = '" & id1 & "'"
                Else
                    sql = "insert into all2016_all (NO1, CABANG, DIVISI, " & _
                        "NO_KONTRAK, NAMA_PROYEK, KODE_PROYEK_LAMA, " & _
                        "KODE_PROYEK_BARU, tgl_bp" & tahun & ", no_bp" & tahun & ", " & _
                        "dpp_pph" & tahun & ") values ('" & _
                        NO1 & "','" & CABANG & "','" & DIVISI & "','" & _
                        NO_KONTRAK & "','" & nama_proyek & "','" & KODE_PROYEK_LAMA & "','" & _
                        KODE_PROYEK_BARU & "','" & set_tgl_perv(tgl_BP) & "','" & no_BP & "','" & _
                        dpp & "')"
                End If
                If ExecSQL1(cnn, sql) <> 0 Then
                    sql = InputBox("error sql", "", sql)
                    p = MsgBox("Stop Load", vbYesNo)
                    If p = vbYes Then stopLoad = True
                End If
            End If
        Else
            'update
            sql = "update all2016_all set tgl_bp" & tahun & " = '" & set_tgl_perv(tgl_BP) & _
                    "', no_bp" & tahun & "='" & no_BP & "', dpp_pph" & tahun & "='" & _
                    dpp & "' where id1 = '" & t & "'"
            If ExecSQL1(cnn, sql) <> 0 Then
                sql = InputBox("error sql", "", sql)
                p = MsgBox("Stop Load", vbYesNo)
                If p = vbYes Then
                    stopLoad = True
                Else
                    Call dbMySQL_open
                End If
            End If
        End If
        
        c = c + 1
        rs.MoveNext
    Loop
    
    Call pesan2("Selesai sin BP")
    stopLoad = True
    Call updateStopLoad
End Sub

Private Sub cmd_load_Click()
    If Me.ch_sin_master.Value = 1 Then
        Me.disable_Form
        Call sin_master
        Me.Enable_Form
    End If
    
    If Me.ch_sinFP.Value = 1 Then
        Me.disable_Form
        Call sin_FP
        Me.Enable_Form
    End If
    
    If Me.ch_sinBP.Value = 1 Then
        Me.disable_Form
        Call sin_BP
        Me.Enable_Form
    End If
    
    'ch_sinWIP
    If Me.ch_sinWIP.Value = 1 Then
        Me.disable_Form
        Call sin_WIP
        Me.Enable_Form
    End If
End Sub

Private Sub cmd_stopLoad_Click()
    stopLoad = True
    updateStopLoad
End Sub

Private Sub Form_Load()
    Dim sql As String
    
    stopLoad = True
    Call updateStopLoad
    
    'load combo
    sql = "select distinct kode_proyek_lama from all2016_all"
    Call Load_combo(Me.cb_proyek, sql, cnn, True, , 1)
End Sub
