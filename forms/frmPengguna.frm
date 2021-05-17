VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmPengguna 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frame_input 
      Caption         =   " Input "
      Height          =   2175
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   10335
      Begin VB.ComboBox cb_divisi 
         Height          =   315
         Left            =   5280
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox cb_Level 
         Height          =   315
         Left            =   5280
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtPassword2 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtPassword1 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   810
         Width           =   2295
      End
      Begin VB.TextBox txtNamaUser 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   330
         Width           =   2295
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Divisi"
         Height          =   195
         Left            =   4680
         TabIndex        =   21
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Level"
         Height          =   195
         Left            =   4680
         TabIndex        =   20
         Top             =   420
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Password (ketik ulang)"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1410
         Width           =   1605
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama User"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   900
         Width           =   690
      End
      Begin VB.Label lb_id1 
         AutoSize        =   -1  'True
         Caption         =   "xx"
         Height          =   195
         Left            =   9840
         TabIndex        =   16
         Top             =   240
         Width           =   150
      End
   End
   Begin VB.CommandButton cmaBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   9240
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtCari 
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Text            =   "Text1"
      ToolTipText     =   "input dan ENTER"
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdUbah 
      Caption         =   "&Ubah"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame frame_data 
      Caption         =   " Data "
      Height          =   2055
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   10335
      Begin MSDataGridLib.DataGrid DGrid1 
         Height          =   1695
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2990
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cari"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   5850
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Data Pengguna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   10575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmPengguna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsGrid As ADODB.Recordset
Dim v_mode_form As Integer


Private Sub cmaBatal_Click()
    Call mode_form(1)
End Sub

Private Sub cmdSimpan_Click()
    
    If cek_password_sama = False Then Exit Sub
    
    If v_mode_form = 2 Then
        'insert
        If tbPengguna_insert(Me.txtNamaUser, Me.txtPassword1, get_kode_combo(Me.cb_Level, "-"), _
                            get_kode_combo(Me.cb_divisi, "-")) = True Then
            Call mode_form(1)
        Else
           Exit Sub
        End If
    ElseIf v_mode_form = 3 Then
        'update
        If tbPengguna_Update(Me.lb_id1.Caption, Me.txtPassword1) = True Then
            Call mode_form(1)
        Else
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdTambah_Click()
    Call mode_form(2)
End Sub

Private Sub cmdUbah_Click()
    Call mode_form(3)
End Sub

Private Sub Form_Load()
    On Error GoTo er1
    Call dbMySQL_open
    Me.Caption = "Penguna"
    Me.txtCari.Text = ""
    Call mode_form(1)
    Exit Sub
er1:
    MsgBox Err.Description, vbCritical
End Sub


Sub load_grid(Optional sql As String = "")
    Dim jRec As Long
    
    '-nm kolom
    '0: nuser, pwd1, level1
    '3: kodedivisi, nama_divisi
    '---------
    
    If Trim(sql) = "" Then
        sql = "SELECT pengguna.nuser, pengguna.pwd1, pengguna.level1, " & _
                "pengguna.kodedivisi, mdivisi.nama_divisi " & _
                "FROM mdivisi RIGHT JOIN pengguna ON mdivisi.kodedivisi = pengguna.kodedivisi " & _
                "order by pengguna.nuser"
    End If
    
    If OpenRecordSet(cnn, rsGrid, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        MsgBox "error run " & sql, vbCritical
        Exit Sub
    End If
    
    jRec = RecordCount(rsGrid)
    If jRec <= 0 Then
        Call pesan2("tidak ada data")
        Exit Sub
    End If
    
    Set Me.DGrid1.DataSource = rsGrid

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set Me.DGrid1.DataSource = Nothing
    Call dbMySQL_close
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
    Dim sql As String
    If KeyAscii = 13 Then
        sql = "SELECT pengguna.nuser, pengguna.pwd1, pengguna.level1, " & _
                "pengguna.kodedivisi, mdivisi.nama_divisi " & _
                "FROM mdivisi RIGHT JOIN pengguna ON mdivisi.kodedivisi = pengguna.kodedivisi " & _
                "where pengguna.nuser like '%" & Trim(Me.txtCari) & "%' or " & _
                "mdivisi.nama_divisi like '%" & Trim(Me.txtCari) & "%' " & _
                "order by pengguna.nuser"
        Call load_grid(sql)
    End If
    
End Sub

Sub mode_form(Optional mode1 As Integer = 1)
    '1: browse
    '2: insert
    '3: update
    
    If mode1 = 2 Then
        v_mode_form = 2
        Me.frame_input.Enabled = Not False
        Me.frame_data.Enabled = Not True
        
        Me.txtNamaUser.Enabled = True
        Me.txtPassword1.Enabled = True
        Me.txtPassword2.Enabled = True
        Me.cb_Level.Enabled = True
        Me.cb_divisi.Enabled = True
        Me.txtCari.Enabled = True
        
        'prepare data
        Me.txtNamaUser.Text = ""
        Me.txtPassword1.Text = ""
        Me.txtPassword2.Text = ""
        
            '1 : admin
            '2 : operator gedung
            '3 : UKP
        Me.cb_Level.Clear
        Me.cb_Level.AddItem "2. Operator Divisi"
        Me.cb_Level.AddItem "3. UKP"
        Me.cb_Level.AddItem "1. Admin"
        Call load_Divisi(Me.cb_divisi, False, 1)
        '------------
        
        Me.cmdTambah.Enabled = Not True
        Me.cmdUbah.Enabled = Not True
        Me.cmdHapus.Enabled = Not True
        Me.cmdSimpan.Enabled = Not False
        Me.cmaBatal.Enabled = Not False
        
        'fokus
        Me.txtNamaUser.SetFocus
    ElseIf mode1 = 3 Then
        v_mode_form = 3
        Me.frame_input.Enabled = Not False
        Me.frame_data.Enabled = Not True
        
        '-nm kolom
        '0: nuser, pwd1, level1
        '3: kodedivisi, nama_divisi
        '---------
        
        'fetch data
        Me.lb_id1.Caption = cek_null(rsGrid(0))
        Me.txtNamaUser.Text = cek_null(rsGrid(0))
        Me.txtPassword1.Text = cek_null(rsGrid(1))
        Me.txtPassword2.Text = cek_null(rsGrid(1))
        Me.cb_Level.Text = cek_null(rsGrid(2))
        Me.cb_divisi.Text = cek_null(rsGrid(3))
        '-----------------
        
        Me.txtNamaUser.Enabled = Not True
        Me.txtPassword1.Enabled = True
        Me.txtPassword2.Enabled = True
        Me.cb_Level.Enabled = Not True
        Me.cb_divisi.Enabled = Not True
        Me.txtCari.Enabled = True
        
        Me.txtCari.Enabled = Not True
        
        Me.cmdTambah.Enabled = Not True
        Me.cmdUbah.Enabled = Not True
        Me.cmdHapus.Enabled = Not True
        Me.cmdSimpan.Enabled = Not False
        Me.cmaBatal.Enabled = Not False
        
        'fokus
        Me.txtPassword1.SetFocus
    Else
        v_mode_form = 1
        Me.frame_input.Enabled = False
        Me.frame_data.Enabled = True
        
        Me.txtNamaUser.Enabled = Not True
        Me.txtPassword1.Enabled = Not True
        Me.txtPassword2.Enabled = Not True
        Me.cb_Level.Enabled = Not True
        Me.cb_divisi.Enabled = Not True
        Me.txtCari.Enabled = True
        
        'clear
        Me.txtNamaUser.Text = ""
        Me.txtPassword1.Text = ""
        Me.txtPassword2.Text = ""
        Me.cb_Level.Clear
        Me.cb_divisi.Clear
        '------------
        
        Me.cmdTambah.Enabled = True
        Me.cmdUbah.Enabled = True
        Me.cmdHapus.Enabled = True
        Me.cmdSimpan.Enabled = False
        Me.cmaBatal.Enabled = False
        
        Call load_grid
    End If

End Sub

Private Sub cmdHapus_Click()
    Dim key1 As String, key2 As String
    Dim p
    Dim jRec As Long
    
    jRec = RecordCount(rsGrid)
    If jRec <= 0 Then
        Call pesan2("tidak ada data")
        Exit Sub
    End If
    
    
    key1 = cek_null(rsGrid(0))
    key2 = cek_null(rsGrid(2))
    p = MsgBox("Yakin menghapus " & key1 & "/" & key2 & "?", vbYesNo)
    If p = vbNo Then
        Call pesan2("dibatalkan")
        Exit Sub
    Else
        If tbPengguna_Delete(key1) = True Then
            Call load_grid
        Else
           Call pesan2("Hapus data gagal")
           Exit Sub
        End If
    End If
    
    
End Sub


Function cek_password_sama() As Boolean
    If Trim(Me.txtPassword1) = (Me.txtPassword2) Then
        cek_password_sama = True
    Else
        cek_password_sama = False
        Call pesan2("Inputan Password tidak sama", , vbYellow)
    End If
End Function

Private Sub txtPassword2_LostFocus()
    Call cek_password_sama
End Sub
