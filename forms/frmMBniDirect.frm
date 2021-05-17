VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmMBniDirect 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCari 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Text            =   "Text1"
      ToolTipText     =   "input dan ENTER"
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   9240
      TabIndex        =   2
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdUbah 
      Caption         =   "&Ubah"
      Height          =   375
      Left            =   8160
      TabIndex        =   1
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Frame frame_data 
      Caption         =   " Data "
      Height          =   4815
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   10335
      Begin MSDataGridLib.DataGrid DGrid1 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7858
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
      TabIndex        =   6
      Top             =   5850
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Master BNI Direct"
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
Attribute VB_Name = "frmMBniDirect"
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


Private Sub cmdTambah_Click()
    Call mode_form(2)
End Sub

Private Sub cmdUbah_Click()
    '-- nmKOlom
    '0: ubni_direct.id1, ubni_direct.kodedivisi, mdivisi.nama_divisi, " & _
    '3: ubni_direct.norek, ubni_direct.nmpemegang, ubni_direct.user_bni
    '----------
    
    Dim id1 As String, kodeDivisi As String, nama_Divisi As String
    Dim noRek As String, nmPemegang As String, user1 As String
    Dim p
    
    If RecordCount(rsGrid) <= 0 Then
        Call pesan2("Tidak ada data", , vbYellow)
        Exit Sub
    End If
    
    id1 = cek_null(rsGrid(0))
    kodeDivisi = cek_null(rsGrid(1))
    nama_Divisi = cek_null(rsGrid(2))
    noRek = cek_null(rsGrid(3))
    nmPemegang = cek_null(rsGrid(4))
    user1 = cek_null(rsGrid(5))
    
    noRek = cleanStr(InputBox("noRek", "Update", noRek))
    nmPemegang = cleanStr(InputBox("nmPemegang", "Update", nmPemegang))
    user1 = cleanStr(InputBox("user1", "Update", user1))
    
    p = MsgBox("Konfirmasi : Ubah Data ?", vbYesNo)
    If p = vbYes Then
        If tbBniDirect_update(kodeDivisi, noRek, nmPemegang, user1) = True Then
                                    
            Call pesan2("Update data Sukses", , vbYellow)
        Else
            Call pesan2("Update data ERROR", , vbYellow)
        End If
    Else
        Call pesan2("Batal", 1, vbYellow)
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo er1
    Call dbMySQL_open
    Me.Caption = "Master User BNI Direct"
    Me.txtCari.Text = ""
    Call mode_form(1)
    Exit Sub
er1:
    MsgBox Err.Description, vbCritical
End Sub


Sub load_grid(Optional sql As String = "")
    Dim jRec As Long, c As Integer
    
    If Trim(sql) = "" Then
        sql = "Select ubni_direct.id1, ubni_direct.kodedivisi, mdivisi.nama_divisi, " & _
                "ubni_direct.norek, ubni_direct.nmpemegang, ubni_direct.user_bni " & _
                "From ubni_direct Left Join " & _
                "mdivisi On mdivisi.kodedivisi = ubni_direct.kodedivisi " & _
                "order By ubni_direct.kodeDivisi"
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
    
    '-- nmKOlom
    '0: ubni_direct.id1, ubni_direct.kodedivisi, mdivisi.nama_divisi, " & _
    '3: ubni_direct.norek, ubni_direct.nmpemegang, ubni_direct.user_bni
    '----------
    
    'format grid
    For c = 0 To rsGrid.Fields.Count - 1
        Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
        
        'lebar kecil
        'If c = 0 Then
        '    Me.DataGrid1.Columns(0).Width = 200
        'End If
    
        'sempit
        If c = 4 Or c = 6 Then
            Me.DGrid1.Columns(c).Width = 1800
        End If
    
        If c = 3 Or c = 7 Then
            Me.DGrid1.Columns(c).Alignment = dbgCenter
            'Me.DGrid1.Columns(c).NumberFormat = "dd mmm yy"
            Me.DGrid1.Columns(c).Width = 1900
        End If
    
        'If c = 6 Then
        '    Me.DataGrid1.Columns(c).Alignment = dbgRight
        '    Me.DataGrid1.Columns(c).NumberFormat = "###,###"
        '    Me.DataGrid1.Columns(c).Width = 1400
        'End If
    Next

End Sub



Private Sub mnImpKpp_Click()
    frm_impMKpp.Show
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
    Dim sql As String
    If KeyAscii = 13 Then
        sql = "Select ubni_direct.id1, ubni_direct.kodedivisi, mdivisi.nama_divisi, " & _
                "ubni_direct.norek, ubni_direct.nmpemegang, ubni_direct.user_bni " & _
                "From ubni_direct Left Join " & _
                "mdivisi On mdivisi.kodedivisi = ubni_direct.kodedivisi " & _
                "where ubni_direct.kodedivisi like '%" & Trim(Me.txtCari) & "%' " & _
                "or mdivisi.nama_divisi like '%" & Trim(Me.txtCari) & "%' " & _
                "or ubni_direct.norek like '%" & Trim(Me.txtCari) & "%' " & _
                "or ubni_direct.nmpemegang like '%" & Trim(Me.txtCari) & "%' " & _
                "or ubni_direct.user_bni like '%" & Trim(Me.txtCari) & "%' " & _
                "order By ubni_direct.kodeDivisi"
        Call load_grid(sql)
    End If
    
End Sub

Sub mode_form(Optional mode1 As Integer = 1)
    '1: browse
    '2: insert
    '3: update
    
    If mode1 = 2 Then
        v_mode_form = 2
        Me.frame_data.Enabled = Not True
        
        Me.cmdUbah.Enabled = Not True
        Me.cmdHapus.Enabled = Not True
        
    ElseIf mode1 = 3 Then
        v_mode_form = 3
        Me.frame_data.Enabled = Not True
        
        '-- nmKOlom
        '0: kodedivisi, nama_divisi, ket
        '----------
        
        'fetch data
        Me.txtCari.Enabled = Not True
        
        Me.cmdUbah.Enabled = Not True
        Me.cmdHapus.Enabled = Not True
        
    Else
        v_mode_form = 1
        Me.frame_data.Enabled = True
        
        
        Me.cmdUbah.Enabled = True
        Me.cmdHapus.Enabled = True
        
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
        If tbBniDirect_delete(key1) = True Then
            Call load_grid
        Else
           Call pesan2("Hapus data gagal")
           Exit Sub
        End If
    End If
    
    
End Sub


