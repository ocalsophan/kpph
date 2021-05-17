VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_repEkualisasi 
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
   Begin VB.CommandButton cmd_stop_Load 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Stop Load"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   " Filter "
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   11895
      Begin VB.ComboBox cbJenis 
         Height          =   330
         Left            =   1200
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox cbTahun 
         Height          =   330
         Left            =   1200
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Jenis"
         Height          =   210
         Left            =   360
         TabIndex        =   7
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun"
         Height          =   210
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   450
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
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
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
      Caption         =   "Report Ekualisasi PPh"
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
Attribute VB_Name = "frm_repEkualisasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sudahLoad As Boolean
Dim stopLoad1 As Boolean

Sub startLoad()
    stopLoad1 = False
    Me.cmd_stop_Load.Visible = True
End Sub

Sub stopLoad()
    stopLoad1 = True
    Me.cmd_stop_Load.Visible = False
End Sub


Sub disable_Form()
    Me.Frame1.Enabled = False
    Me.cmd_proses.Enabled = False
    'Me.Frame2.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame1.Enabled = True
    Me.cmd_proses.Enabled = True
    'Me.Frame2.Enabled = True
End Sub

Private Sub cmd_proses_Click()
    Dim s As String
    Dim p
    
    On Error GoTo er1
    
    Me.disable_Form
    
    
    If Trim(Me.cbJenis) = "" Then
        Call pesan2("Jenis Report belum dipilih", , vbYellow)
        Me.Enable_Form
        Exit Sub
    End If
    
    Call dbMySQL_open
    Call create_ds_Access("c:\dbpph.dsn", App.Path & "\data\", App.Path & "\data\dbrep.mdb")
    
    If sudahLoad = False Then
        p = MsgBox("ReLoad data trial balance ke temporary ?" & vbCr & _
                    "Proses ini dilakukan jika ada TrialBalance Baru", vbYesNo)
        Me.disable_Form
        Call startLoad
        If p = vbYes Then Call fetch_dbRep_tbAccpac(Me.cbTahun.text, Me.StatusBar1, stopLoad1)
        Call stopLoad
        
        
        p = MsgBox("Lakukan Fetch Data ?", vbYesNo)
        If p = vbYes Then
            Call startLoad
            Call fetch_dbRep_tbAccpac_subkon(Me.cbTahun.text, Me.StatusBar1, stopLoad1)
            Call startLoad
            Call fetch_dbRep_tbAccpac_22(Me.cbTahun.text, Me.StatusBar1, stopLoad1)
            Call startLoad
            Call fetch_dbRep_tbAccpac_23(Me.cbTahun.text, Me.StatusBar1, stopLoad1)
            Call startLoad
            Call fetch_dbRep_tbAccpac_21(Me.cbTahun.text, Me.StatusBar1, stopLoad1)
            Call stopLoad
        End If
    End If
    
    sudahLoad = True
    
    If Trim(Me.cbJenis.text) = "Subkon" Then
        s = "{tbaccpac_subkon.tahun} = '" & Me.cbTahun.text & "'"
        Call tampil_report(CR, App.Path & "\rep\rep_ekual_subkon.rpt", 85, s)
    ElseIf Trim(Me.cbJenis.text) = "23" Then
        s = "{tbaccpac_23.tahun} = '" & Me.cbTahun.text & "'"
        Call tampil_report(CR, App.Path & "\rep\rep_ekual_23.rpt", 85, s)
    ElseIf Trim(Me.cbJenis.text) = "22" Then
        s = "{tbaccpac_22.tahun} = '" & Me.cbTahun.text & "'"
        Call tampil_report(CR, App.Path & "\rep\rep_ekual_22.rpt", 85, s)
    ElseIf Trim(Me.cbJenis.text) = "21" Then
        s = "{tbaccpac_21.tahun} = '" & Me.cbTahun.text & "'"
        Call tampil_report(CR, App.Path & "\rep\rep_ekual_21.rpt", 85, s)
    Else
        Call pesan2("Jenis Report tidak valid", , vbYellow)
    End If
    
    
    
    Me.Enable_Form
    Exit Sub
er1:
    MsgBox Err.Description, vbCritical
    Me.Enable_Form
End Sub

Private Sub cmd_stop_Load_Click()
    Call stopLoad
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
  Call load_TahunEkualisasi(Me.cbTahun)
  Me.cbTahun.ListIndex = Me.cbTahun.ListCount - 1
  
  Me.cbJenis.Clear
  Me.cbJenis.AddItem "Subkon"
  Me.cbJenis.AddItem "22"
  Me.cbJenis.AddItem "23"
  Me.cbJenis.AddItem "21"
  Me.cbJenis.ListIndex = 0
  
  sudahLoad = False
  Call stopLoad
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Kill "C:\dbpph.dsn"
    Call dbMySQL_close
End Sub


