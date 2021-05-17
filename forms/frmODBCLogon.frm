VERSION 5.00
Begin VB.Form frmODBCLogon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ODBC Logon"
   ClientHeight    =   1650
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   4470
   ControlBox      =   0   'False
   Icon            =   "frmODBCLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   450
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   450
      Left            =   2880
      TabIndex        =   3
      Top             =   1080
      Width           =   1440
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Save"
      Height          =   450
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   1440
   End
   Begin VB.Frame fraStep3 
      Caption         =   "Connection Values"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4230
      Begin VB.TextBox txtUID 
         Height          =   300
         Left            =   1125
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "IP Server"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   0
         Top             =   390
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmODBCLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call lokasi_server_save(App.Path & "\data\set_db.txt", Me.txtUID)
End Sub

Private Sub cmdTest_Click()
    If dbMySQL_open = True Then
        MsgBox "Open Database Sukses", vbInformation
    Else
        MsgBox "Open Database Tidak Sukses", vbExclamation
    End If
End Sub

Private Sub Form_Load()
    txtUID = lokasi_server_load(App.Path & "\data\set_db.txt")
End Sub
