VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   1785
   ClientLeft      =   2790
   ClientTop       =   3150
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1054.637
   ScaleMode       =   0  'User
   ScaleWidth      =   5506.917
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   3330
      TabIndex        =   1
      Top             =   285
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   4140
      TabIndex        =   5
      Top             =   1170
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3330
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   675
      Width           =   2325
   End
   Begin VB.Image imgLogo 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1425
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   2145
      TabIndex        =   0
      Top             =   300
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   2145
      TabIndex        =   2
      Top             =   690
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
  On Error Resume Next
    
    If dbMySQL_open = True Then
        'MsgBox "Open Database Sukses", vbInformation
        
    Else
        MsgBox "Open Database Tidak Sukses", vbExclamation
    End If
  
    'check for correct password
    If Trim(Me.txtUserName) = "admin" And Trim(Me.txtPassword) = "admin2013" Then
        LoginSucceeded = True
        frMenu1.nmLogin = UCase(Me.txtUserName)
        Call frMenu1.tampil_Menu(1)
        Unload Me
    ElseIf tbPengguna_isValid_Password(Me.txtUserName, Me.txtPassword) > 0 Then
        LoginSucceeded = True
        
        frMenu1.nmLogin = (Me.txtUserName)
        Call frMenu1.tampil_Menu(tbPengguna_getLevel1(Me.txtUserName))
        
        Unload Me
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
    Call frMenu1.set_Caption(True)
    
End Sub


Sub set_Enable(mode1 As Boolean)
    Me.txtUserName.Enabled = mode1
    Me.txtPassword.Enabled = mode1
    Me.cmdOK.Enabled = mode1
    Me.cmdCancel.Enabled = mode1
End Sub

Private Sub Form_Load()

    Dim a As Long, hasil As Long

    Call Me.set_Enable(False)
    frMenu1.set_Caption
    Call frMenu1.tampil_Menu(False)
    
    If dbMySQL_open = True Then
        'MsgBox "Open Database Sukses", vbInformation
        
    Else
        MsgBox "Open Database Tidak Sukses", vbExclamation
    End If
    
    'Me.txtUserName = "ukp3"
    'Me.txtPassword = "1234"
    
    'Me.txtUserName = "adm1"
    'Me.txtPassword = "a"
    
    'Me.txtUserName = "infra1"
    'Me.txtPassword = "ppcab-2"
    
    Call Me.set_Enable(True)
    
End Sub



Private Sub txtPassword_GotFocus()
    Call selAllText(Me.txtPassword)
End Sub

Private Sub txtUserName_GotFocus()
    Call selAllText(Me.txtUserName)
End Sub
