VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_repSSP_PPh 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3255
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
   ScaleHeight     =   3255
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
      TabIndex        =   15
      Top             =   2160
      Width           =   7335
      Begin VB.OptionButton opt_detil 
         Caption         =   "Detil"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opt_rekap 
         Caption         =   "Rekap"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport CR 
      Left            =   11040
      Top             =   840
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
      TabIndex        =   8
      Top             =   2280
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
      TabIndex        =   10
      Top             =   600
      Width           =   12015
      Begin VB.ComboBox cb_pph_ssp 
         Height          =   330
         Left            =   1320
         TabIndex        =   3
         Text            =   "x"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.ComboBox cb_masa 
         Height          =   330
         Left            =   8160
         TabIndex        =   5
         Text            =   "x"
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cb_tahun 
         Height          =   330
         Left            =   8160
         TabIndex        =   4
         Text            =   "x"
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cb_KPP 
         Height          =   330
         Left            =   1320
         TabIndex        =   2
         Text            =   "Combo1"
         ToolTipText     =   "F2 untuk Filter"
         Top             =   720
         Width           =   5415
      End
      Begin VB.ComboBox cb_divisi 
         Height          =   330
         Left            =   1320
         TabIndex        =   1
         Text            =   "x"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jenis PPh SSP"
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   1140
         Width           =   1035
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
         TabIndex        =   14
         Top             =   780
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   7560
         TabIndex        =   13
         Top             =   420
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "KPP"
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   780
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Divisi"
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   420
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3000
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
      Caption         =   "Report SSP PPh"
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
Attribute VB_Name = "frm_repSSP_PPh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function cek_Isian() As Boolean
    Dim pesan1 As String, t As String
    Dim hasil As Boolean
    
    pesan1 = ""
    hasil = True
    
    'cek divisi
    If Trim(Me.cb_divisi.Text) = "" Then
        hasil = False
        pesan1 = pesan1 & "Divisi tidak valid"
    End If
    
    'cek KPP
    If Trim(Me.cb_KPP.Text) = "" Then
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
    If KeyCode = 113 And Shift = 0 Then Call load_KPP(Me.cb_KPP, True)
End Sub




Sub disable_Form()
    Me.Frame1.Enabled = False
    Me.Frame2.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame1.Enabled = True
    Me.Frame2.Enabled = True
End Sub

Private Sub cmd_proses_Click()
    Dim jenisPPh As String
    
    On Error GoTo er1
    
    Me.disable_Form
    
    If cek_Isian() = False Then
        Me.Enable_Form
        Exit Sub
    End If
    
    Call dbMySQL_open
    Call create_ds_Access("c:\dbpph.dsn", App.Path & "\data\", App.Path & "\data\dbrep.mdb")
    
    Call fetch_dbRep_Divisi(get_kode_combo(Me.cb_divisi, "-"), Me.StatusBar1)
    Call fetch_dbRep_KPP(get_kode_combo(Me.cb_KPP, "#"), Me.StatusBar1)
    
    Call fetch_dbRep_PPhX(get_kode_combo(Me.cb_KPP, "#"), get_kode_combo(Me.cb_divisi, "-"), "", Me.cb_tahun.Text, _
                                Me.cb_masa.Text, Me.StatusBar1, "ssp_pph", Me.cb_pph_ssp)
                                
    If Me.opt_detil.Value = True Then
        Call tampil_report(CR, App.Path & "\rep\repSSP_Pph.rpt", 85)
    Else
        Call tampil_report(CR, App.Path & "\rep\repSSP_Pph_rekap.rpt", 85)
    End If
    
    Me.Enable_Form
    Exit Sub
er1:
    MsgBox Err.Description, vbCritical
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
  Call load_KPP(Me.cb_KPP, False, 1)
  Call load_jenisPPhSsp(Me.cb_pph_ssp)
  
  Call load_Tahun2(Me.cb_tahun, "ssp_pph")
  Call load_Masa2(Me.cb_masa, "ssp_pph")
  
    'get level
  '2 : operator gedung
  '3 : UKP
  
  '---------
  
  Level1 = tbPengguna_getLevel1(frMenu1.nmLogin)
  If Level1 = 2 Then
    Me.cb_divisi.Text = tbPengguna_getDivisi(frMenu1.nmLogin)
    Me.cb_divisi.Enabled = False
  ElseIf Level1 = 3 Then
    Me.cb_divisi.Enabled = True
  Else
    Call pesan2("Level tidak valid", , vbYellow)
    Me.cb_divisi.Enabled = False
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Kill "C:\dbpph.dsn"
    Call dbMySQL_close
End Sub


