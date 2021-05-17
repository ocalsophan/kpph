VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form LOV_2 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Of Valid Values"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   ControlBox      =   0   'False
   FillColor       =   &H80000013&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   4815
      Begin MSDataGridLib.DataGrid grid1 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   -2147483629
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         WrapCellPointer =   -1  'True
         RowDividerStyle =   5
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Batal"
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   4320
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdFind 
         Appearance      =   0  'Flat
         Caption         =   "&Cari"
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtKeyword 
         Height          =   350
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Kata Kunci :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "=="
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   5760
      Width           =   180
   End
End
Attribute VB_Name = "LOV_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BY KAUTSAR 21 JUN 05
'WITH DATAGRID
'updated : 20 Agustus 2010

Option Explicit

Dim rs As ADODB.Recordset

Sub load_grid(sql As String)
  Dim a As Integer
  Dim lebar As Integer
  Dim jml_record As Integer
  
  
  
  If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) = 0 Then
    
    jml_record = RecordCount(rs)
    Me.Label2 = "Jumlah Data:" & jml_record
    If jml_record > 0 Then
      Set grid1.DataSource = rs
      If rs.Fields.Count < 2 Then
        MsgBox "Tidak ada deskripsi", vbInformation, "Load Grid"
        grid1.Columns(0).Caption = rs.Fields(0).Name
      End If
      
      For a = 1 To rs.Fields.Count
        'header
        grid1.Columns(a - 1).Caption = UCase(rs.Fields(a - 1).Name)
        'lebar kolom
        If a - 1 = 0 Then
          grid1.Columns(a - 1).Width = lov_width_field1
        ElseIf a - 1 = 1 Then
          grid1.Columns(a - 1).Width = lov_width_field2
        Else
          lebar = Len(rs.Fields(a - 1).Name)
          If rs.Fields(a - 1).ActualSize > lebar Then
            lebar = rs.Fields(a - 1).ActualSize
          End If
          grid1.Columns(a - 1).Width = lebar * 150
        End If
        'visible kolom ke 5 dst
        If a > 3 Then grid1.Columns(a - 1).Visible = False
      Next
      grid1.Columns(0).Alignment = dbgCenter
    End If
  Else
    MsgBox "Error open tabel.", vbCritical, "Load Grid"
  End If
  Exit Sub
End Sub

Private Sub cmdCancel_Click()
  On Error Resume Next
  'lov_return = ""
  Set rs = Nothing
  Unload Me
End Sub

Private Sub CmdOk_Click()
  On Error Resume Next
  lov_return = rs.Fields(0).Value
  Set grid1.DataSource = Nothing
  Set rs = Nothing
  Call Form_Unload(1)
End Sub

Private Sub cmdFind_Click()
  Dim sql As String
  
  sql = susun_SqL
  Call load_grid(sql)
  'MsgBox sql
End Sub


Function susun_SqL() As String
  Dim sql2 As String
  Dim where2
  Dim a As Integer
  
  'menyusun sql
  'Public lov_sql_Kolom : isi SQL: select kd_grup, nm_group form gvendor
  'Public lov_kolom_Dicari, pisah dgn koma : kd_grup, nm_grup
  'Public lov_order_by: nama kolom order by
  
  sql2 = lov_SqL
  If Trim(Me.txtKeyword) <> "" Then
    where2 = Split(lov_kolom_Dicari, ",")
    For a = 0 To UBound(where2, 1)
      If a = 0 Then
        sql2 = sql2 & " where ucase(" & where2(a) & ") like '%" & UCase(Trim(Me.txtKeyword)) & "%' "
      Else
        sql2 = sql2 & " or ucase(" & where2(a) & ") like '%" & UCase(Trim(Me.txtKeyword)) & "%' "
      End If
    Next
  End If
  
  If Trim(lov_order_by) <> "" Then
    sql2 = sql2 & " order by " & Trim(lov_order_by)
  End If
  susun_SqL = sql2
End Function

Private Sub Form_Load()
  Dim sql As String
  
  On Error Resume Next
  Me.Caption = lov_title
  
  
  Me.txtKeyword = lov_Key_Cari
  sql = susun_SqL
  
  load_grid (sql)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub

Private Sub grid1_DblClick()
  Call CmdOk_Click
End Sub

Private Sub grid1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call cmdCancel_Click
  End If
End Sub

Private Sub txtKeyword_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call cmdFind_Click
  ElseIf KeyAscii = 27 Then
    Call cmdCancel_Click
  End If
End Sub
