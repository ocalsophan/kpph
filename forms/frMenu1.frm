VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frMenu1 
   BackColor       =   &H8000000C&
   Caption         =   "App CSV"
   ClientHeight    =   4875
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7140
   Icon            =   "frMenu1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4620
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6006
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6006
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   960
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mnLogin 
      Caption         =   "Login"
   End
   Begin VB.Menu mnDataLogin 
      Caption         =   "Data Login"
   End
   Begin VB.Menu mnData 
      Caption         =   "Data"
      Begin VB.Menu mnMDivisi 
         Caption         =   "Master Divisi"
      End
      Begin VB.Menu mnMKpp 
         Caption         =   "Master KPP"
      End
      Begin VB.Menu mnMaster_Ptkp 
         Caption         =   "Master PTKP"
      End
      Begin VB.Menu g1b 
         Caption         =   "-"
      End
      Begin VB.Menu mnBniDirect 
         Caption         =   "Mandiri Direct"
      End
      Begin VB.Menu mnMKppCode 
         Caption         =   "KPP Code"
      End
      Begin VB.Menu g1c 
         Caption         =   "-"
      End
      Begin VB.Menu mnVariabel 
         Caption         =   "Variabel"
      End
   End
   Begin VB.Menu mnTransfer 
      Caption         =   "Transfer Data"
      Begin VB.Menu mnLoadCsv 
         Caption         =   "Import CSV SPT dari Divisi"
      End
      Begin VB.Menu mnImpCsvSSP 
         Caption         =   "Import CSV SSP dari Divisi"
      End
      Begin VB.Menu g1 
         Caption         =   "-"
      End
      Begin VB.Menu mnMNpwp 
         Caption         =   "Master NPWP"
         Visible         =   0   'False
      End
      Begin VB.Menu mnMasterKaryawan 
         Caption         =   "Master Karyawan"
      End
      Begin VB.Menu g2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImpPPh21Tahunan2 
         Caption         =   "Import PPh 21 Tahunan(2)"
      End
      Begin VB.Menu mnImpPPh21Tahunan2_edit 
         Caption         =   "Import PPh 21 Tahunan(2) - edit by ID"
      End
      Begin VB.Menu g3 
         Caption         =   "-"
      End
      Begin VB.Menu mnImptbAccpac 
         Caption         =   "Import TrialBalance Accpac"
      End
      Begin VB.Menu mnSAPPPh 
         Caption         =   "SAP PPh"
         Begin VB.Menu mnImpSapPPh 
            Caption         =   "Import"
         End
         Begin VB.Menu mnRepSapPPh 
            Caption         =   "Report"
         End
      End
      Begin VB.Menu mnEbupot23 
         Caption         =   "eBuPot23"
         Begin VB.Menu mnAlur 
            Caption         =   "Alur"
         End
         Begin VB.Menu mnEbupotPPh23 
            Caption         =   "PPh23"
         End
         Begin VB.Menu mnEbupot26 
            Caption         =   "PPh26"
         End
         Begin VB.Menu mnExp_eBupot23 
            Caption         =   "Export eBupot23"
         End
         Begin VB.Menu mnImpBuktiPotongDJP 
            Caption         =   "Import Bukti Potong (dari DJP)"
         End
      End
   End
   Begin VB.Menu mnLaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnRepSptPPh 
         Caption         =   "Data SPT PPh"
      End
      Begin VB.Menu mnRepSSSPpph 
         Caption         =   "Data SSP PPh"
      End
      Begin VB.Menu mnRepKaryawan 
         Caption         =   "Karyawan"
      End
      Begin VB.Menu mnBrowse 
         Caption         =   "Browse Data"
      End
      Begin VB.Menu mnDashboar 
         Caption         =   "Dashboard SPT / SSP"
      End
      Begin VB.Menu g01 
         Caption         =   "-"
      End
      Begin VB.Menu mnCekData 
         Caption         =   "Cek Data PPh21Tahunan(2)"
      End
      Begin VB.Menu mnDashPPh21Th 
         Caption         =   "Dashboard PPh21 Tahunan"
      End
      Begin VB.Menu mnBrowsePPhTahunan2 
         Caption         =   "Browse Data PPh 21 Tahunan (2)"
      End
      Begin VB.Menu mnRekPPh21Tahunan 
         Caption         =   "Rekap PPh 21 Tahunan"
      End
      Begin VB.Menu mnBuktiPotong 
         Caption         =   "Bukti Potong "
      End
      Begin VB.Menu mnRepRincianBeban 
         Caption         =   "Rekap data PPh21Tahunan -  Per Proyek"
      End
      Begin VB.Menu g02 
         Caption         =   "-"
      End
      Begin VB.Menu mnBrwDtTrialBalance 
         Caption         =   "Browse Data Trial Balance"
      End
      Begin VB.Menu mnEqualisasi 
         Caption         =   "Ekualisasi"
         Begin VB.Menu mnEkPPh 
            Caption         =   "PPh Masa"
         End
         Begin VB.Menu mnEkuMasProyek 
            Caption         =   "Master Proyek"
         End
         Begin VB.Menu mn_mas_acc 
            Caption         =   "COA"
         End
         Begin VB.Menu mnEkImport 
            Caption         =   "Data"
            Begin VB.Menu mnEkImport_TB 
               Caption         =   "Trial Balance"
            End
            Begin VB.Menu mnEkImport_fp 
               Caption         =   "FP - Keluaran / Masukan"
            End
            Begin VB.Menu mnEkImport_bp 
               Caption         =   "BP"
            End
            Begin VB.Menu mnEkGab 
               Caption         =   "Gabungan"
            End
            Begin VB.Menu mng1 
               Caption         =   "-"
            End
            Begin VB.Menu mnEkPPh2 
               Caption         =   "Ekualisasi PPh"
            End
            Begin VB.Menu mnEkPPN2 
               Caption         =   "Ekualisasi PPN"
            End
         End
         Begin VB.Menu mnEkPPN 
            Caption         =   "PPN ALL 2016"
            Begin VB.Menu mnEkPpnDataLengkap 
               Caption         =   "Data Lengkap"
            End
         End
         Begin VB.Menu mnRekapPPh 
            Caption         =   "Rekap PPh - ekualisasi"
         End
      End
      Begin VB.Menu mnSptBadan 
         Caption         =   "SPT Badan"
      End
   End
   Begin VB.Menu mnExpSPT 
      Caption         =   "exp CSV"
      Begin VB.Menu mnExpCsvSpt1 
         Caption         =   "CSV SPT"
      End
      Begin VB.Menu mnExpCsvSspPph 
         Caption         =   "CSV SSP PPh"
      End
      Begin VB.Menu mnExpCsvBniDirect 
         Caption         =   "Mandiri MFT"
      End
   End
   Begin VB.Menu mnLain 
      Caption         =   "Lain"
      Begin VB.Menu mnKphWeb 
         Caption         =   "Kpph WEB"
      End
      Begin VB.Menu mnSetKoneksi 
         Caption         =   "Set Koneksi"
      End
      Begin VB.Menu mnTentang 
         Caption         =   "Tentang"
      End
      Begin VB.Menu mnTest 
         Caption         =   "test"
      End
   End
End
Attribute VB_Name = "frMenu1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nmLogin As String

Sub tampil_Menu(mode1 As Integer)
    '1 : admin
    '2 : operator gedung
    '3 : UKP
    
    If mode1 = 1 Then
        'admin
        Me.mnLogin.Visible = True
        Me.mnDataLogin.Visible = Not False
        Me.mnData.Visible = False
        Me.mnTransfer.Visible = False
        Me.mnLaporan.Visible = False
        Me.mnImptbAccpac.Visible = False
        mnExpSPT.Visible = False
    ElseIf mode1 = 2 Then
        'operator gedung
        Me.mnLogin.Visible = True
        Me.mnDataLogin.Visible = False
        Me.mnData.Visible = False
        Me.mnTransfer.Visible = Not False
        Me.mnLaporan.Visible = Not False
        Me.mnImptbAccpac.Visible = False
        mnExpSPT.Visible = True
        mnExpCsvSpt1.Visible = False
        mnExpCsvSspPph.Visible = False
        mnExpCsvBniDirect.Visible = True
    ElseIf mode1 = 3 Then
        'operator UKP
        Me.mnLogin.Visible = True
        Me.mnDataLogin.Visible = False
        Me.mnData.Visible = Not False
        Me.mnTransfer.Visible = Not False
        Me.mnLaporan.Visible = Not False
        Me.mnImptbAccpac.Visible = True
        mnExpSPT.Visible = True
        mnExpCsvSpt1.Visible = True
        mnExpCsvSspPph.Visible = True
        mnExpCsvBniDirect.Visible = True
    Else
        'not valid
        Me.mnLogin.Visible = True
        Me.mnDataLogin.Visible = False
        Me.mnData.Visible = False
        Me.mnTransfer.Visible = False
        Me.mnLaporan.Visible = False
        Me.mnImptbAccpac.Visible = False
        mnExpSPT.Visible = False
    End If
End Sub

Sub set_Caption(Optional setLevel As Boolean = False)
    Dim t As String
    t = "db" & App.title & ". v " & App.Major & "." & App.Minor & "." & App.Revision
    Me.Caption = t & " @" & lokasi_server_load(App.Path & "\data\set_db.txt")
     
    Call info(1, "Login: " & Me.nmLogin, Me.StatusBar1)
    If setLevel = True Then
        Call dbMySQL_open
        Call info(2, "Level: " & tbPengguna_getLevel1(Me.nmLogin), Me.StatusBar1)
        Call dbMySQL_close
    End If
End Sub

Private Sub MDIForm_Load()
    frmSplash.Show vbModal
    '-- cek versi app
    If cek_Lng(get_versiapp()) >= cek_Lng(tbVariabel_get("versikpph")) Then
    Else
        Call pesan2("Versi Aplikasi KPPH harus di update", , vbYellow)
    End If
    Call update_tabel_temp
    Call set_Caption
    Call tampil_Menu(0)
End Sub

Private Sub mn_mas_acc_Click()
    frm_EkMas_Acc.Show
End Sub

Private Sub mnAlur_Click()
    Dim alur As String
    
    alur = "1. Operator DVO upload PPh23/26 sesuai template dari pusat" & vbCr & _
           "2. Sistem akan cek kesesuaian kode, npwp, nik" & vbCr & _
           "3. pusat download berdasarkan masa dan KPP" & vbCr & _
           "4. file siap di upload ke DJP, dari DJP mendapatkan no Bukti Potong" & vbCr & _
           "5. noBukti potong di import ke Kpph, merge dengan data SPT pph23/26"
    MsgBox alur, vbInformation
End Sub

Private Sub mnBniDirect_Click()
    frmMBniDirect.Show
End Sub

Private Sub mnBrowse_Click()
    frm_Browse.Show
End Sub

Private Sub mnBrowsePPhTahunan2_Click()
    frm_Browse_PPh21Tahunan2.Show
End Sub

Private Sub mnBrwDtTrialBalance_Click()
    frm_Browse_TB.Show
End Sub

Private Sub mnBuktiPotong_Click()
    frm_BuktiPotong.Show
End Sub

Private Sub mnCekData_Click()
    frm_Browse_CekData.Show
End Sub

Private Sub mnDashboar_Click()
    frm_Dashboard.Show
End Sub

Private Sub mnDashPPh21Th_Click()
    frm_Dashboard_pph21.Show
End Sub

Private Sub mnDataLogin_Click()
    frmPengguna.Show
End Sub

Private Sub mnEbupot26_Click()
    frm_Ebupotpph26.Show
End Sub

Private Sub mnEbupotPPh23_Click()
    frm_Ebupotpph23.Show
End Sub

Private Sub mnEkGab_Click()
    frm_Ek_All.Show
End Sub

Private Sub mnEkImport_bp_Click()
    frm_EkBP.Show
End Sub

Private Sub mnEkImport_fp_Click()
    frm_EkFP.Show
End Sub

Private Sub mnEkImport_TB_Click()
    frm_Ektb.Show
End Sub

Private Sub mnEkPph_Click()
    Dim p
    
    p = MsgBox("Format ekualisasi ini sudah tidak dipakai. " & vbCr & _
            "Tetap Lanjut ? ", vbYesNo)
    If p = vbYes Then frm_repEkualisasi.Show
End Sub

Private Sub mnEkPPh2_Click()
    frm_Ek_PPh.Show
End Sub

Private Sub mnEkPPN2_Click()
    frm_Ek_PPn.Show
End Sub

Private Sub mnEkPpnDataLengkap_Click()
    frm_EkPpnAll.Show
End Sub

Private Sub mnEkuMasProyek_Click()
    frm_EkMastProyek.Show
End Sub

Private Sub mnExp_eBupot23_Click()
    frm_Ebupotpph23_ALL.Show
End Sub

Private Sub mnExpCsvBniDirect_Click()
    frm_csv_BNI.Show
End Sub

Private Sub mnExpCsvSpt1_Click()
    frm_csv_SPT.Show
End Sub

Private Sub mnExpCsvSspPph_Click()
    frm_csv_Ssp.Show
End Sub

Private Sub mnImpBuktiPotongDJP_Click()
    frm_Ebupot23_res.Show
End Sub

Private Sub mnImpCsvSSP_Click()
    frm_impCSVssp.Show
End Sub

Private Sub mnImpPPh21Tahunan2_Click()
    frm_impPph21Tahunan2.Show
End Sub

Private Sub mnImpPPh21Tahunan2_edit_Click()
    frm_impPph21Tahunan2_edit.Show
End Sub

Private Sub mnImpSapPPh_Click()
    frm_SAP_PPh_Imp.Show
End Sub

Private Sub mnImptbAccpac_Click()
    frm_impTbAccpac.Show
End Sub

Private Sub mnKphWeb_Click()
    Dim url1 As String
    
    On Error GoTo er1
    url1 = "explorer.exe " & "http://acc.ptpp.co.id/kpph"
    'File1 = InputBox("", "", File1)
    Call Shell(url1, vbNormalFocus)
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION
End Sub

Private Sub mnLoadCsv_Click()
    frm_impCSV.Show
End Sub

Private Sub mnLogin_Click()
    frmLogin.Show vbModal
End Sub

Private Sub mnMaster_Ptkp_Click()
    frmM_Ptkp.Show
End Sub

Private Sub mnMasterKaryawan_Click()
    frm_mKaryawan.Show
End Sub

Private Sub mnMDivisi_Click()
    frmMaster_Divisi.Show
End Sub

Private Sub mnMKpp_Click()
    frmMKpp.Show
End Sub

Private Sub mnMKppCode_Click()
    frmMKppCode.Show
End Sub

Private Sub mnMNpwp_Click()
    frm_mNpwp.Show
End Sub

Private Sub mnRekapPPh_Click()
    frm_RekapPPhEk.Show
End Sub

Private Sub mnRekPPh21Tahunan_Click()
    Call dbMySQL_open
    Call info(2, "cek biaya_jabatan", frMenu1.StatusBar1)
    Call tbPph21Tahunan2_setBiayaJabatanPerBulan
    Call info(2, "cek biaya_jabatan : OK", frMenu1.StatusBar1)
    
    frm_Rekap_PPh21Tahunan2.Show
End Sub

Private Sub mnRepKaryawan_Click()
    frm_repKaryawan.Show
End Sub

Private Sub mnRepRincianBeban_Click()
    frm_RepRincianBebanPPh21T2.Show
End Sub

Private Sub mnRepSapPPh_Click()
    frm_SAP_PPh_Browse.Show
End Sub

Private Sub mnRepSptPPh_Click()
    frm_repPPh.Show
End Sub

Private Sub mnRepSSSPpph_Click()
    frm_repSSP_PPh.Show
End Sub

Private Sub mnSetKoneksi_Click()
    frmODBCLogon.Show
End Sub

Private Sub mnSptBadan_Click()
    frm_SPTBadan.Show
End Sub

Private Sub mnTentang_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnTest_Click()
    Dim res As String, fileWinscp As String
    Dim d1 As Date, d2 As Date
    Dim isecords As Long
    
     '-- cek apakah file winscp sudah ada ?
        fileWinscp = "c:\Program Files (x86)\WinSCP\WinSCP.com"
        If Dir(fileWinscp) <> "" Then
            MsgBox "File exists"
        Else
            MsgBox "File does not exist"
        End If
        '------
    
    
    'd1 = Now
    'd2 = DateAdd("s", -30, d1)
    'isecords = DateDiff("s", d1, d2)
    
    'Call tbVariabel_set("test", Now)
    'res = tbVariabel_get("test")
    'MsgBox res
    'd1 = res
    'd2 = Now
    'isecords = DateDiff("s", d1, d2)
    'MsgBox s
    
    'MsgBox (Format(d1, "ddmmyy hhmmss") & "-" & Format(d2, "ddmmyy hhmmss") & "-" & isecords)
    'res = EncryptString(cleanStr("test"), "dvak2017")
    'MsgBox res
    'MsgBox checkNPWP("914073770454000")
    'Call DownloadFile("ftp.trunojoyopython.com", "user3@trunojoyopython.com", "programmer2019")
    'Call UploadFile("182.253.5.26", "pp", "PpPajak2019", "21")
End Sub

Private Sub mnVariabel_Click()
    frmTVariabel.Show
End Sub
