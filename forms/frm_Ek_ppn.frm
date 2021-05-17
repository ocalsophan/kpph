VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Ek_ppn 
   ClientHeight    =   7245
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
   ScaleHeight     =   7245
   ScaleWidth      =   12300
   Begin VB.ListBox List1 
      Height          =   5100
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   12015
   End
   Begin VB.Frame Frame1 
      Caption         =   " Filter "
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   12015
      Begin VB.TextBox txtTahun 
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   390
         Width           =   855
      End
      Begin VB.CommandButton cmd_load 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Load"
         Height          =   375
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   240
         TabIndex        =   3
         Top             =   442
         Width           =   450
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
      Caption         =   "Ekualisasi : PPN"
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
Attribute VB_Name = "frm_Ek_PPn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nama_data As String



Private Sub cmd_load_Click()
    Dim nmFile As String, File1 As String, fileSimpan As String, tahun As String
    Dim fl As Object
    Dim baris As Integer
    Dim sql As String, t As String, tahunLalu As String
    
    Dim prestasi_YAMP As Currency, karya_YDF As Currency, WIP2 As Currency
    Dim a11103 As Currency, a11102 As Currency, nilai As Currency
    Dim a11601 As Currency, a21201 As Currency
    
    Me.List1.Clear
    
    Me.List1.AddItem "Opening DB"
    Call dbMySQL_open
    
    tahun = Trim(Me.txtTahun)
    If Trim(tahun) = "" Then Exit Sub
    tahunLalu = CStr(CInt(tahun) - 1)
    Me.List1.AddItem "Tahun: " & tahun & "- tahun Lalu:" & tahunLalu
    nmFile = App.Path & "\rep\200825EqualissiPPN.xlsx"
    
        
    'open file xls
    Me.List1.AddItem "open xls"
    If open_xls_lateBinding(fl, nmFile) <> 0 Then
        Call pesan2("error open template")
        close_xls_lateBinding (fl)
    End If
    
    'tulis
    Me.List1.AddItem "nilai akun 40101 s/d 40511"
    baris = 4
    fl.Cells(baris, 23).Value = "Januari s.d. Desember " & tahun
    baris = 12
    sql = "SELECT `F_get_nilaiAkunAllBetween`('40101', '40511', '" & tahun & "') "
    t = cari_data1(cnn, sql, True)
    fl.Cells(baris, 17).Value = t
    
    Me.List1.AddItem "get DPP PPN"
    sql = "SELECT `F_get_DppPPNAll`('" & tahun & "') "
    t = cari_data1(cnn, sql, True)
    fl.Cells(baris, 22).Value = t
    
    'Uang muka pelanggan akhir
    Me.List1.AddItem "UM pelanggan"
    baris = 14
    sql = "SELECT `F_get_nilaiAkunAll`('20501','" & tahun & "') "
    t = cari_data1(cnn, sql, True)
    fl.Cells(baris, 17).Value = t
    
    'Penyerahan tahun sebelumnya difakturkan tahun ini
    'wip2'=11601-21201
    Me.List1.AddItem "WIP2 tahun Lalu"
    baris = 19
    sql = "SELECT `F_get_nilaiAkunAll`('11601','" & tahunLalu & "') "
    prestasi_YAMP = cek_Money(cari_data1(cnn, sql, True))
    sql = "SELECT `F_get_nilaiAkunAll`('21201','" & tahunLalu & "') "
    karya_YDF = cek_Money(cari_data1(cnn, sql, True))
    WIP2 = prestasi_YAMP - karya_YDF
    fl.Cells(baris, 17).Value = WIP2
    
    'Uang muka pelanggan awal
    Me.List1.AddItem "UM pelanggan awal"
    baris = 27
    sql = "SELECT `F_get_nilaiAkunAll`('20501','" & tahunLalu & "') "
    t = cari_data1(cnn, sql, True)
    fl.Cells(baris, 17).Value = t
    
    'Piutang Retensi
    Me.List1.AddItem "piutang retensi"
    baris = 29
    sql = "SELECT `F_get_nilaiAkunAll`('11102','" & tahun & "') "
    a11102 = cek_Money(cari_data1(cnn, sql, True))
    sql = "SELECT `F_get_nilaiAkunAll`('11103','" & tahun & "') "
    a11103 = cek_Money(cari_data1(cnn, sql, True))
    nilai = a11102 + a11102
    fl.Cells(baris, 17).Value = nilai
    
    'Penyerahan difakturkan tahun berikutnya
    Me.List1.AddItem "penyerahan difakturkan"
    baris = 30
    sql = "SELECT `F_get_nilaiAkunAll`('11601','" & tahun & "') "
    a11601 = cek_Money(cari_data1(cnn, sql, True))
    sql = "SELECT `F_get_nilaiAkunAll`('21201','" & tahun & "') "
    a21201 = cek_Money(cari_data1(cnn, sql, True))
    nilai = a11601 + 21201
    fl.Cells(baris, 17).Value = nilai
    
    'Jumlah Penyerahan non BKP/JKP
    Me.List1.AddItem "jumlah penyerahan nonBKP/JKP"
    baris = 34
    sql = "SELECT `F_get_nilaiAkunAll`('40102','" & tahun & "') "
    t = cari_data1(cnn, sql, True)
    fl.Cells(baris, 17).Value = t
    
    '---- simpan
    Me.List1.AddItem "saving.."
    'fl.ActiveWorkbook.Save
    fileSimpan = App.Path & "\exp\EqualissiPPN" & tahun & ".xlsx"
    fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
    'open by explorer
    File1 = "explorer.exe " & fileSimpan
    Call Shell(File1, vbNormalFocus)

End Sub

Private Sub Form_Load()
    Me.txtTahun.text = Year(Now)
    
    
    Me.Height = 7710
    Me.Width = 12420
End Sub
