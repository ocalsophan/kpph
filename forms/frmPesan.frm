VERSION 5.00
Begin VB.Form frmPesan 
   BackColor       =   &H008080FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   780
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   2010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H80000005&
   ForeColor       =   &H80000005&
   Icon            =   "frmPesan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   2010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer buka 
      Left            =   120
      Top             =   360
   End
   Begin VB.Timer tutup 
      Left            =   1440
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   105
   End
End
Attribute VB_Name = "frmPesan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tinggi As Long

Private Sub buka_Timer()
  Me.Height = Me.Height + 50
  
  Me.Top = Round(Screen.Height / 2) - Me.Height
  
  If Me.Height >= tinggi Then
    Me.buka.Enabled = False
    Me.Timer1.Enabled = True
  End If
End Sub

Private Sub Form_Activate()
  tinggi = Me.Height
  Me.Height = 0
  
  Me.buka.Interval = 50
  Me.buka.Enabled = True
End Sub

Private Sub Form_Load()
  Me.tutup.Enabled = False
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub

Private Sub Timer1_Timer()
  'Unload Me
  Me.Timer1.Enabled = False
  Me.tutup.Interval = 50
  Me.tutup.Enabled = True
End Sub


Private Sub tutup_Timer()
  Me.Height = Me.Height - 50
  Me.Top = Me.Top + 50
  If Me.Height <= 50 Then
    Me.Enabled = False
    Unload Me
  End If
End Sub
