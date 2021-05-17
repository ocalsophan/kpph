VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmMaster_Divisi 
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
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   10335
      Begin VB.TextBox txtKeterangan 
         Height          =   735
         Left            =   5520
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtNamaDivisi 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtKode 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
         Height          =   195
         Left            =   4560
         TabIndex        =   17
         Top             =   450
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Kode Divisi"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   450
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nama Divisi"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   930
         Width           =   840
      End
      Begin VB.Label lb_id1 
         AutoSize        =   -1  'True
         Caption         =   "xx"
         Height          =   195
         Left            =   9840
         TabIndex        =   14
         Top             =   240
         Width           =   150
      End
   End
   Begin VB.CommandButton cmaBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   9240
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtCari 
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Text            =   "Text1"
      ToolTipText     =   "input dan ENTER"
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdUbah 
      Caption         =   "&Ubah"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "&Tambah"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Frame frame_data 
      Caption         =   " Data "
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   10335
      Begin MSDataGridLib.DataGrid DGrid1 
         Height          =   2535
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   4471
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
      TabIndex        =   12
      Top             =   5850
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Master Divisi"
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
Attribute VB_Name = "frmMaster_Divisi"
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
    
    Call dbMySQL_open
    If v_mode_form = 2 Then
        'insert
        If tbMDivisi_insert(Me.txtKode, Me.txtNamaDivisi, Me.txtKeterangan) = True Then
            Call mode_form(1)
        Else
           Exit Sub
        End If
    ElseIf v_mode_form = 3 Then
        'update
        If tbMDivisi_Update(Me.lb_id1.Caption, Me.txtNamaDivisi, Me.txtKeterangan) = True Then
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
    
    Me.Caption = "Master Divisi"
    Me.txtCari.Text = ""
    Call mode_form(1)
    
    Exit Sub
er1:
    MsgBox Err.Description, vbCritical
End Sub


Sub load_grid(Optional sql As String = "")
    Dim jRec As Long
    
    '-- nmKOlom
    '0: kodedivisi, nama_divisi, ket
    '----------
    
    If Trim(sql) = "" Then
        sql = "select kodedivisi, nama_divisi, ket from mdivisi order by kodedivisi"
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
    Call dbMySQL_close
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
    Dim sql As String
    If KeyAscii = 13 Then
            Call dbMySQL_open
        sql = "select kodedivisi, nama_divisi, ket from mdivisi " & _
                "where kodedivisi like '%" & Trim(Me.txtCari) & "%' or " & _
                "nama_divisi like '%" & Trim(Me.txtCari) & "%' order by kodedivisi"
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
        
        Me.txtKode.Enabled = True
        Me.txtNamaDivisi.Enabled = True
        Me.txtKeterangan.Enabled = True
        Me.txtCari.Enabled = True
        
        'prepare data
        Me.txtKode.Text = ""
        Me.txtNamaDivisi.Text = ""
        Me.txtKeterangan.Text = ""
        '------------
        
        Me.cmdTambah.Enabled = Not True
        Me.cmdUbah.Enabled = Not True
        Me.cmdHapus.Enabled = Not True
        Me.cmdSimpan.Enabled = Not False
        Me.cmaBatal.Enabled = Not False
        
        'fokus
        Me.txtKode.SetFocus
    ElseIf mode1 = 3 Then
        v_mode_form = 3
        Me.frame_input.Enabled = Not False
        Me.frame_data.Enabled = Not True
        
        '-- nmKOlom
        '0: kodedivisi, nama_divisi, ket
        '----------
        
        'fetch data
        Me.lb_id1.Caption = cek_null(rsGrid(0))
        Me.txtKode.Text = cek_null(rsGrid(0))
        Me.txtNamaDivisi = cek_null(rsGrid(1))
        Me.txtKeterangan = cek_null(rsGrid(2))
        '-----------------
        Me.txtKode.Enabled = Not True
        Me.txtNamaDivisi.Enabled = True
        Me.txtKeterangan.Enabled = True
        Me.txtCari.Enabled = Not True
        
        Me.cmdTambah.Enabled = Not True
        Me.cmdUbah.Enabled = Not True
        Me.cmdHapus.Enabled = Not True
        Me.cmdSimpan.Enabled = Not False
        Me.cmaBatal.Enabled = Not False
        
        'fokus
        Me.txtNamaDivisi.SetFocus
    Else
        v_mode_form = 1
        Me.frame_input.Enabled = False
        Me.frame_data.Enabled = True
        
        'clear
        Me.txtKode.Text = ""
        Me.txtNamaDivisi.Text = ""
        Me.txtKeterangan.Text = ""
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
    key2 = cek_null(rsGrid(1))
    p = MsgBox("Yakin menghapus " & key1 & "/" & key2 & "?", vbYesNo)
    If p = vbNo Then
        Call pesan2("dibatalkan")
        Exit Sub
    Else
        If tbMDivisi_Delete(key1) = True Then
            Call load_grid
        Else
           Call pesan2("Hapus data gagal")
           Exit Sub
        End If
    End If
    
    
End Sub


