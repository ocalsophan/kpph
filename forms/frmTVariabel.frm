VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmTVariabel 
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
      TabIndex        =   2
      Text            =   "Text1"
      ToolTipText     =   "input dan ENTER"
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdUbah 
      Caption         =   "&Ubah"
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Frame frame_data 
      Caption         =   " Data "
      Height          =   4815
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   10335
      Begin MSDataGridLib.DataGrid DGrid1 
         Height          =   4455
         Left            =   120
         TabIndex        =   4
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
      TabIndex        =   5
      Top             =   5880
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Master Variabel"
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
Attribute VB_Name = "frmTVariabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsGrid As ADODB.Recordset
Dim v_mode_form As Integer
Dim nama_data As String

Private Sub cmaBatal_Click()
    Call mode_form(1)
End Sub


Private Sub cmdTambah_Click()
    Call mode_form(2)
End Sub

Private Sub cmdUbah_Click()
    Dim val1(11)
    Dim a As Integer
    Dim p
    Dim sql As String
    Dim start_kolom As Integer
    
    start_kolom = 2
    If RecordCount(rsGrid) <= 0 Then Exit Sub
    
    p = MsgBox("Ubah data " & nama_data & "?", vbYesNo)
    If p = vbNo Then Exit Sub
        
        
    For a = start_kolom To rsGrid.Fields.Count - 1
        val1(a) = InputBox(rsGrid.Fields(a).Name, "Input", cek_null(rsGrid.Fields(a).Value))
    Next
    
    sql = "update tvariabel set "
    For a = start_kolom To rsGrid.Fields.Count - 1
        If a = rsGrid.Fields.Count - 1 Then
            sql = sql & rsGrid.Fields(a).Name & "='" & val1(a) & "'"
        Else
            sql = sql & rsGrid.Fields(a).Name & "='" & val1(a) & "', "
        End If
    Next
    sql = sql & " where `id1` = '" & rsGrid.Fields(0).Value & "'"
    
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("sql error", "", sql)
    Else
        Call load_grid
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo er1
    Call dbMySQL_open
    Me.Caption = "Master Variabel"
    nama_data = "Master Variabel"
    Me.txtCari.text = ""
    Call mode_form(1)
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
End Sub


Sub load_grid(Optional sql As String = "")
    Dim jRec As Long, c As Integer
    
    If Trim(sql) = "" Then
        sql = "select id1, key1, ket " & _
                "from tvariabel order by id1"
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
    '0: npwp, nama, alamat
    '3: tgl_lahir, klu, nip_nama_ar
    '6: status_update, tgl_update, kpp_administrasi
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
            Me.DGrid1.Columns(c).Width = 800
        End If
    
        If c = 3 Or c = 7 Then
            Me.DGrid1.Columns(c).Alignment = dbgCenter
            Me.DGrid1.Columns(c).NumberFormat = "dd mmm yy"
            Me.DGrid1.Columns(c).Width = 900
        End If
    
        'If c = 6 Then
        '    Me.DataGrid1.Columns(c).Alignment = dbgRight
        '    Me.DataGrid1.Columns(c).NumberFormat = "###,###"
        '    Me.DataGrid1.Columns(c).Width = 1400
        'End If
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call dbMySQL_close
End Sub

Private Sub mnImpKpp_Click()
    frm_impMKpp.Show
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
    Dim sql As String
    If KeyAscii = 13 Then
        sql = "select id1, key1, ket " & _
                "from tvariabel " & _
                "where key1 like '%" & Trim(Me.txtCari) & "%' " & _
                "or ket like '%" & Trim(Me.txtCari) & "%' order by id1"
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
        
    ElseIf mode1 = 3 Then
        v_mode_form = 3
        Me.frame_data.Enabled = Not True
        
        '-- nmKOlom
        '0: kodedivisi, nama_divisi, ket
        '----------
        
        'fetch data
        Me.txtCari.Enabled = Not True
        
        Me.cmdUbah.Enabled = Not True
        
    Else
        v_mode_form = 1
        Me.frame_data.Enabled = True
        
        
        Me.cmdUbah.Enabled = True
    
        
        Call load_grid
    End If

End Sub


