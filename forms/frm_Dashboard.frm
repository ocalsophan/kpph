VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Dashboard 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8700
   ClientLeft      =   225
   ClientTop       =   735
   ClientWidth     =   11655
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
   ScaleHeight     =   8700
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   3000
      Index           =   4
      Left            =   3960
      TabIndex        =   12
      ToolTipText     =   "Double Klik untuk menyimpan"
      Top             =   5160
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Height          =   3000
      Index           =   3
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Double Klik untuk menyimpan"
      Top             =   5160
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Index           =   2
      Left            =   7800
      TabIndex        =   10
      ToolTipText     =   "Double Klik untuk menyimpan"
      Top             =   2160
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Index           =   1
      Left            =   3960
      TabIndex        =   9
      ToolTipText     =   "Double Klik untuk menyimpan"
      Top             =   2160
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Index           =   0
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Double Klik untuk menyimpan"
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   9720
      Top             =   840
   End
   Begin VB.Frame Frame1 
      Caption         =   " Filter Data "
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4095
      Begin VB.CommandButton cmd_togle 
         BackColor       =   &H00C0FFC0&
         Caption         =   "x"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "togle refresh"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtMasa 
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtTahun 
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Masa"
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   915
         Width           =   390
      End
      Begin VB.Label Label1 
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
      Top             =   8445
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10239
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10239
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
   Begin VB.Label lb_info 
      Alignment       =   1  'Right Justify
      Caption         =   "xx"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   7800
      TabIndex        =   13
      Top             =   5160
      Width           =   3735
   End
   Begin VB.Label lb_caption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dashboard Data SSP - SPT PPh"
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
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11565
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frm_Dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub disable_Form()
    Dim c As Integer
    
    Me.Frame1.Enabled = False
    For c = 0 To 4
        Me.List1(c).Enabled = False
    Next
End Sub

Sub Enable_Form()
    Dim c As Integer
    
    Me.Frame1.Enabled = Not False
    For c = 0 To 4
        Me.List1(c).Enabled = Not False
    Next
End Sub


Private Sub cmd_togle_Click()
    Dim t As Long
    
    On Error GoTo er1
    If Timer1.Enabled = False Then
        Me.disable_Form
        t = CLng(InputBox("input (detik)", "Interval Refresh (dalam detik)?", "30"))
        Me.Timer1.Interval = t * 1000
        Call dbMySQL_open
        Call Load_Data
        Me.Enable_Form
    End If
    Me.Timer1.Enabled = Not Me.Timer1.Enabled
    Call status_tombol
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
End Sub

Private Sub Form_Load()
    Me.txtTahun.text = Year(Now)
    Me.txtMasa.text = Month(DateAdd("m", -1, Now))
    Me.Timer1.Enabled = False
    Call status_tombol
End Sub

Sub status_tombol()
    If Me.Timer1.Enabled = True Then
        Me.cmd_togle.Caption = "Auto Refresh"
    Else
        Me.cmd_togle.Caption = "Refresh OFF"
    End If
End Sub

Sub saveLog()
    Dim pesan
    Dim namaFile As String, t1 As String
    Dim f
    Dim idx As Integer, c As Integer
  
  pesan = MsgBox("Simpan File Log ? ", vbYesNo)
  If pesan = vbYes Then
    Me.disable_Form
    Me.Timer1.Enabled = False
    namaFile = "d:\LogDashboard-" & Year(Date) & "_" & Month(Date) & "_" & Day(Date) & " _ " & _
               "j" & Hour(Time) & Minute(Time) & Second(Time) & ".txt"
    Call OpenFile(namaFile, f, 2)
    For c = 0 To 4
        For idx = 0 To List1(c).ListCount - 1
            Me.List1(c).ListIndex = idx
            t1 = List1(c).text & Chr(13) & Chr(10)
            Call writefile(f, t1)
        Next
    Next
    Call closefile(f)
    MsgBox "File export di simpan di " & namaFile, vbInformation
    Me.Enable_Form
    Me.Timer1.Enabled = True
  End If
End Sub

Private Sub List1_DblClick(Index As Integer)
    Call saveLog
End Sub

Private Sub Timer1_Timer()
    Me.Timer1.Enabled = False
    Me.disable_Form
    Call dbMySQL_open
    Call Load_Data
    Me.Enable_Form
    Me.Timer1.Enabled = True
End Sub

Sub Load_Data()
    Dim c As Integer
    
    For c = 0 To 4
        Me.List1(c).Clear
    Next
    
    Call load_ssp
End Sub

Function load_ssp() As String
    Dim sql As String, t As String
    Dim jenis(4) As String, dvo(5) As String
    Dim cPPh As Integer, cDvo As Integer
    Dim jumlahTemp As Currency, grandTotal As Currency
    Dim jumlahPPh(4) As Currency
    
    'get jenis PPH
    jenis(0) = "PPh Pasal 22"
    jenis(1) = "PPh Pasal 23"
    jenis(2) = "PPh Final"
    jenis(3) = "PPh Pasal 21"
    jenis(4) = "PPh Pasal 15"
        
    'get dvo
    dvo(0) = "100000"
    dvo(1) = "200000"
    dvo(2) = "300000"
    dvo(3) = "400000"
    dvo(4) = "500000"
    dvo(5) = "700000"
    
    'get total pph per DVO
    grandTotal = 0
    For cPPh = 0 To UBound(jenis)
        Me.List1(cPPh).AddItem "*** " & jenis(cPPh) & " ***"
        jumlahPPh(cPPh) = 0
        For cDvo = 0 To UBound(dvo)
            Me.List1(cPPh).AddItem "- Divisi " & dvo(cDvo) & " -"
            sql = "select sum(Jumlah_SSP) from ssp_pph where Jenis_Pajak = '" & jenis(cPPh) & _
                    "' and kode_divisi = '" & dvo(cDvo) & "' and Tahun_Pajak = '" & _
                    Trim(Me.txtTahun) & "' and (Masa_Pajak = '" & Trim(Me.txtMasa) & "' or Masa_Pajak = '" & adddigit(CLng(Trim(Me.txtMasa)), 2) & "') "
            t = cari_data1(cnn, sql, True)
            Me.List1(cPPh).AddItem " . Jumlah SSP: " & Format(CCur(t), "###,###")
            
            
            '---- get data SPT
            If jenis(cPPh) = "PPh Pasal 22" Then
                
                sql = "select sum(Nilai_PPh) from pph22 where kode_divisi = '" & dvo(cDvo) & "' and Tahun_Pajak = '" & _
                    Trim(Me.txtTahun) & "' and (Masa_Pajak = '" & Trim(Me.txtMasa) & "' or Masa_Pajak = '" & adddigit(CLng(Trim(Me.txtMasa)), 2) & "') "
                t = cari_data1(cnn, sql, True)
                Me.List1(cPPh).AddItem " . Jumlah SPT PPh: " & Format(CCur(t), "###,###")
                
                
                grandTotal = grandTotal + CCur(t)
                jumlahPPh(cPPh) = jumlahPPh(cPPh) + CCur(t)
                Call info(1, "total Rp." & Format(grandTotal, "###,###"), Me.StatusBar1)
                
            ElseIf jenis(cPPh) = "PPh Pasal 23" Then
                sql = "select sum(Jumlah_PPh_Yang_Dipotong) from pph23 where kode_divisi = '" & dvo(cDvo) & "' and Tahun_Pajak = '" & _
                    Trim(Me.txtTahun) & "' and (Masa_Pajak = '" & Trim(Me.txtMasa) & "' or Masa_Pajak = '" & adddigit(CLng(Trim(Me.txtMasa)), 2) & "') "
                t = cari_data1(cnn, sql, True)
                Me.List1(cPPh).AddItem " . Jumlah SPT PPh: " & Format(CCur(t), "###,###")
                
                grandTotal = grandTotal + CCur(t)
                jumlahPPh(cPPh) = jumlahPPh(cPPh) + CCur(t)
                Call info(1, "total Rp." & Format(grandTotal, "###,###"), Me.StatusBar1)
            ElseIf jenis(cPPh) = "PPh Final" Then
            
                sql = "select sum(PPh_Yang_Dipotong__1 + PPh_Yang_Dipotong__2 + PPh_Yang_Dipotong__3 + " & _
                        "PPh_Yang_Dipotong__4 + PPh_Yang_Dipotong__5 + PPh_Yang_Dipotong__6 " & _
                        ") from pph42_konstruksi " & _
                        "where kode_divisi = '" & dvo(cDvo) & "' and Tahun_Pajak = '" & _
                        Trim(Me.txtTahun) & "' and " & _
                        "(Masa_Pajak = '" & Trim(Me.txtMasa) & "' or Masa_Pajak = '" & adddigit(CLng(Trim(Me.txtMasa)), 2) & "') " & _
                        "and right(kode_form,3) <> '317'"
                t = cari_data1(cnn, sql, True)
                jumlahTemp = CCur(t)
                
                sql = "select sum(Jumlah_PPh_Yang_Dipotong) from pph42_konstruksi " & _
                        "where kode_divisi = '" & dvo(cDvo) & "' and Tahun_Pajak = '" & _
                        Trim(Me.txtTahun) & "' and " & _
                        "(Masa_Pajak = '" & Trim(Me.txtMasa) & "' or Masa_Pajak = '" & adddigit(CLng(Trim(Me.txtMasa)), 2) & "') " & _
                        "and right(kode_form,3) = '317'"
                t = cari_data1(cnn, sql, True)
                
                jumlahTemp = jumlahTemp + CCur(t)
                
                
                sql = "select sum(PPh_Yang_Dipotong__1 + PPh_Yang_Dipotong__2 + PPh_Yang_Dipotong__3 + " & _
                        "PPh_Yang_Dipotong__4 + PPh_Yang_Dipotong__5 + PPh_Yang_Dipotong__6  " & _
                        ") from pph42_sewa " & _
                        "where kode_divisi = '" & dvo(cDvo) & "' and Tahun_Pajak = '" & _
                        Trim(Me.txtTahun) & "' and " & _
                        "(Masa_Pajak = '" & Trim(Me.txtMasa) & "' or Masa_Pajak = '" & adddigit(CLng(Trim(Me.txtMasa)), 2) & "') " & _
                        "and right(kode_form,3) <> '317'"
                t = cari_data1(cnn, sql, True)
                jumlahTemp = jumlahTemp + CCur(t)
                
                sql = "select sum(Jumlah_PPh_Yang_Dipotong) from pph42_sewa " & _
                        "where kode_divisi = '" & dvo(cDvo) & "' and Tahun_Pajak = '" & _
                        Trim(Me.txtTahun) & "' and " & _
                        "(Masa_Pajak = '" & Trim(Me.txtMasa) & "' or Masa_Pajak = '" & adddigit(CLng(Trim(Me.txtMasa)), 2) & "') and right(kode_form,3) = '317'"
                t = cari_data1(cnn, sql, True)
                jumlahTemp = jumlahTemp + CCur(t)
                
                '--- pph42_obligasi
                sql = "select sum(PPh_Yang_Dipotong__1 + PPh_Yang_Dipotong__2 + PPh_Yang_Dipotong__3 + " & _
                        "PPh_Yang_Dipotong__4 + PPh_Yang_Dipotong__5 + PPh_Yang_Dipotong__6  " & _
                        ") from pph42_obligasi " & _
                        "where kode_divisi = '" & dvo(cDvo) & "' and Tahun_Pajak = '" & _
                        Trim(Me.txtTahun) & "' and " & _
                        "(Masa_Pajak = '" & Trim(Me.txtMasa) & "' or Masa_Pajak = '" & adddigit(CLng(Trim(Me.txtMasa)), 2) & "') " & _
                        "and right(kode_form,3) <> '317'"
                t = cari_data1(cnn, sql, True)
                jumlahTemp = jumlahTemp + CCur(t)
                
                sql = "select sum(Jumlah_PPh_Yang_Dipotong) from pph42_obligasi " & _
                        "where kode_divisi = '" & dvo(cDvo) & "' and Tahun_Pajak = '" & _
                        Trim(Me.txtTahun) & "' and " & _
                        "(Masa_Pajak = '" & Trim(Me.txtMasa) & "' or Masa_Pajak = '" & adddigit(CLng(Trim(Me.txtMasa)), 2) & "') and right(kode_form,3) = '317'"
                t = cari_data1(cnn, sql, True)
                jumlahTemp = jumlahTemp + CCur(t)
                '-------------
                
                
                Me.List1(cPPh).AddItem " . Jumlah SPT PPh: " & Format(CCur(jumlahTemp), "###,###")
                grandTotal = grandTotal + CCur(jumlahTemp)
                jumlahPPh(cPPh) = jumlahPPh(cPPh) + CCur(jumlahTemp)
                Call info(1, "total Rp." & Format(grandTotal, "###,###"), Me.StatusBar1)
            ElseIf jenis(cPPh) = "PPh Pasal 21" Then
            
                sql = "select sum(Jumlah_PPh) from pph21bulanan where kode_divisi = '" & dvo(cDvo) & "' and Tahun_Pajak = '" & _
                    Trim(Me.txtTahun) & "' and (Masa_Pajak = '" & Trim(Me.txtMasa) & "' or Masa_Pajak = '" & adddigit(CLng(Trim(Me.txtMasa)), 2) & "') "
                t = cari_data1(cnn, sql, True)
                jumlahTemp = CCur(t)
                
                sql = "select sum(Jumlah_PPh) from pph21tf where kode_divisi = '" & dvo(cDvo) & "' and Tahun_Pajak = '" & _
                    Trim(Me.txtTahun) & "' and (Masa_Pajak = '" & Trim(Me.txtMasa) & "' or Masa_Pajak = '" & adddigit(CLng(Trim(Me.txtMasa)), 2) & "') "
                t = cari_data1(cnn, sql, True)
                jumlahTemp = jumlahTemp + CCur(t)
                
                
                Me.List1(cPPh).AddItem " . Jumlah SPT PPh: " & Format(CCur(jumlahTemp), "###,###")
                grandTotal = grandTotal + CCur(jumlahTemp)
                jumlahPPh(cPPh) = jumlahPPh(cPPh) + CCur(jumlahTemp)
                Call info(1, "total Rp." & Format(grandTotal, "###,###"), Me.StatusBar1)
            ElseIf jenis(cPPh) = "PPh Pasal 15" Then
            
                sql = "select sum(pph_dipotong) from pph15 where kode_divisi = '" & dvo(cDvo) & "' and Tahun_Pajak = '" & _
                    Trim(Me.txtTahun) & "' and (Masa_Pajak = '" & Trim(Me.txtMasa) & "' or Masa_Pajak = '" & adddigit(CLng(Trim(Me.txtMasa)), 2) & "') "
                t = cari_data1(cnn, sql, True)
                Me.List1(cPPh).AddItem " . Jumlah SPT PPh: " & Format(CCur(t), "###,###")
                
                grandTotal = grandTotal + CCur(t)
                jumlahPPh(cPPh) = jumlahPPh(cPPh) + CCur(t)
                Call info(1, "total Rp." & Format(grandTotal, "###,###"), Me.StatusBar1)
            End If
            
        Next
        'Me.List1.AddItem ""
    Next
    
    t = ""
    For cPPh = 0 To 4
        t = t & "total SPT " & jenis(cPPh) & ": Rp." & Format(jumlahPPh(cPPh), "###,###") & vbCr
    Next
    
    t = t & ":: total SPT PPH" & vbCr & "Rp." & Format(grandTotal, "###,###") & ",-"
    Me.lb_info.Caption = t
    
    
    
End Function



