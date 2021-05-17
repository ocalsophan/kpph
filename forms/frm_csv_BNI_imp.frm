VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_csv_BNI_imp 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7245
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
   ScaleHeight     =   7245
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   6990
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
   Begin VB.ListBox List1 
      Height          =   1110
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Log Hasil Import File Excel. Double Klik untuk Simpan"
      Top             =   5610
      Width           =   12045
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   " 2. Isi File "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Left            =   117
      TabIndex        =   7
      Top             =   2040
      Width           =   12038
      Begin VB.CommandButton cmd_import 
         Caption         =   "Import ke DB"
         Height          =   375
         Left            =   10560
         TabIndex        =   4
         Top             =   2880
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DGrid1 
         Height          =   2460
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   11790
         _ExtentX        =   20796
         _ExtentY        =   4339
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
   Begin VB.Frame Frame2 
      Caption         =   " 1. Pilih File Import "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   12045
      Begin VB.TextBox txtKarakkter 
         Height          =   375
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmd_browse 
         Caption         =   "Browse"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   10365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Karakter pemisah"
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   915
         Width           =   1260
      End
   End
   Begin VB.Label lb_caption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Import Tax Inquiry Report"
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
Attribute VB_Name = "frm_csv_BNI_imp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset


Private Sub cmd_browse_Click()
  Dim f As String
  Dim jmlKolom As Integer
  
  On Error GoTo er1
  
  '---------------------
    If dbMySQL_open = True Then
    Else
        MsgBox "Open Database Tidak Sukses", vbExclamation
    End If
    '---------------------
  
  
  MsgBox "Salah Pilih Format akan menampilkan hasil yang salah", vbExclamation
  Me.disable_Form
  CD.InitDir = App.Path & "\Import\"
  CD.Filter = "xls / CSV file (*.csv;*.txt;*.xls)|*.csv;*.txt;*.xls"
  CD.FileName = ""
  CD.ShowOpen
  f = CD.FileName
  
  'cek inputan karakter pemisah
  If Trim(Me.txtKarakkter) = "" Then
    Call pesan2("Karakter pemisah tidak boleh kosong", , vbYellow)
    Me.Enable_Form
    Exit Sub
  End If
  '----
  
  If Trim(f) <> "" Then
    Me.Text1 = f
    If is_file_ada(f) = True Then
      'File Valid
        Call Load_Csv_2Rs(f, rs, Me.StatusBar1, Trim(Me.txtKarakkter.text), 0)
        Me.cmd_import.Enabled = True
    Else
      MsgBox "File tidak valid", vbCritical
      Me.cmd_import.Enabled = False
    End If
  End If
  MsgBox "Jumlah data di file : " & RecordCount(rs)
  Set Me.DGrid1.DataSource = rs
  
  'jumlah kolom
  '1 : 15 kolom
  '2 : 51 kolom
  '3 : 19
  '4 : 9
  '5 : 51
  '6 : 77
  '7 : 51
  
  
  If RecordCount(rs) <= 0 Then
    Me.cmd_import.Enabled = False
    Me.Enable_Form
    Exit Sub
  End If
  
  jmlKolom = rs.Fields.Count
  
  Me.cmd_import.Enabled = True
  If jmlKolom = 31 Then
  Else
    Call pesan2("jumlah kolom Tidak Valid", , vbYellow)
    Me.cmd_import.Enabled = False
  End If
  
  
  '------------
  Me.Enable_Form
  Exit Sub
er1:
  MsgBox Err.DESCRIPTION, vbCritical
  Me.Enable_Form
End Sub

Sub disable_Form()
    Me.Frame2.Enabled = False
    Me.Frame3.Enabled = False
    Me.List1.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame2.Enabled = True
    Me.Frame3.Enabled = True
    Me.List1.Enabled = True
End Sub

Private Sub cmd_import_Click()
    Dim jRec As Long, jenisPPh As String
    Dim ps
  
    On Error GoTo er1
    
    'konfirmasi,
    ps = MsgBox("Yakin akan import Data ?" & vbCr & "Pastikan Regional Setting: Indonesia", vbYesNo)
    If ps = vbNo Then Exit Sub
  
    jRec = RecordCount(rs)
    If jRec <= 0 Then
        MsgBox "Tidak ada data"
        Exit Sub
    End If
    Me.disable_Form
    
    
    Call import_data
    Me.Enable_Form
    Exit Sub
er1:
  MsgBox Err.DESCRIPTION, vbCritical
  Me.Enable_Form
End Sub


Sub import_data()
    Dim jRec As Long, c As Long, jml_Insert As Long, jml_Update As Long
  
    Dim NO1 As String, trans_reference_no As String, created_date As String
    Dim transaction_date As String, posting_date As String
    Dim billing_id As String, NTB As String, NTPN As String, STAN As String
    Dim tax_type As String, deposite_type As String, NPWP_number As String
    Dim tax_payer_name As String, City As String, WP_Address As String, NPWP_Payer As String
    Dim Payer_Name As String, Payer_Address As String, NOP As String, Tax_Period As String
    Dim SK_Number As String, Customer_Reference_No As String, Beneficiary_Email As String
    Dim Remark As String, Extended_Payment_Detail As String, Currency1 As String
    Dim Amount As Currency, Signature_ID As String, Signature_Name As String, Status As String
    Dim Reason As String
    
    Dim return1 As Integer
    '--------------------------
    Dim t As String, ps, sql As String
    Dim data_Valid As Boolean
  
    jRec = RecordCount(rs)
    If jRec <= 0 Then
        MsgBox "Tidak ada data"
        Exit Sub
    End If
      
    rs.MoveFirst
    c = 0
    jml_Insert = 0
    jml_Update = 0
    Me.List1.Clear
    Do While rs.EOF = False
        c = c + 1
        Call info(1, "Run Import " & c & " of " & jRec & ". Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update, Me.StatusBar1)
        data_Valid = True
    
        'ssssssssssssssssss
        NO1 = cek_null(rs(0))
        
        If UCase(NO1) = "NO" Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " skip")
        End If
        
        trans_reference_no = cek_null(rs(1))
        If Trim(trans_reference_no) = "" Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " trans_reference_no kosong. skip")
        End If
        
        created_date = cek_null(rs(2))
        transaction_date = cek_null(rs(3))
        posting_date = cek_null(rs(4))
        billing_id = cek_null(rs(5))
        NTB = cek_null(rs(6))
        NTPN = cek_null(rs(7))
        If Trim(NTPN) = "" Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " NTPN Kosong. skip")
        End If
        
        STAN = cek_null(rs(8))
        tax_type = cek_null(rs(9))
        deposite_type = cek_null(rs(10))
        NPWP_number = cek_null(rs(11))
        tax_payer_name = cek_null(rs(12))
        City = cek_null(rs(13))
        WP_Address = cek_null(rs(14))
        NPWP_Payer = cek_null(rs(15))
        Payer_Name = cek_null(rs(16))
        Payer_Address = cek_null(rs(17))
        NOP = cek_null(rs(18))
        Tax_Period = cek_null(rs(19))
        SK_Number = cek_null(rs(20))
        Customer_Reference_No = cek_null(rs(21))
        If Trim(Customer_Reference_No) = "" Then
            data_Valid = False
            Call setListInfo(Me.List1, "Data ke " & c & " Customer_Reference_No Kosong. skip")
        End If
        
        Beneficiary_Email = cek_null(rs(22))
        Remark = cek_null(rs(23))
        Extended_Payment_Detail = cek_null(rs(24))
        Currency1 = cek_null(rs(25))
        Amount = cek_Money(rs(26))
        Signature_ID = cek_null(rs(27))
        Signature_Name = cek_null(rs(28))
        Status = cek_null(rs(29))
        Reason = cek_null(rs(30))
    
        If data_Valid = True Then
        
            return1 = tbTaxInqury_insert(NO1, trans_reference_no, created_date, transaction_date, posting_date, _
                                    billing_id, NTB, NTPN, STAN, tax_type, _
                                    deposite_type, NPWP_number, tax_payer_name, City, WP_Address, _
                                    NPWP_Payer, Payer_Name, Payer_Address, NOP, Tax_Period, _
                                    SK_Number, Customer_Reference_No, Beneficiary_Email, Remark, Extended_Payment_Detail, _
                                    Currency1, Amount, Signature_ID, Signature_Name, Status, _
                                    Reason)
            If return1 = 1 Then
                jml_Insert = jml_Insert + 1
            ElseIf return1 = 2 Then
                jml_Update = jml_Update + 1
            Else
                Call pesan2("Insert error", , vbYellow)
                Exit Sub
            End If
        End If
        rs.MoveNext
    Loop
    MsgBox "Proses Import Selesai. Jml Import: " & jml_Insert & _
                 ". Jml Update: " & jml_Update, vbInformation

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
  
  'get level
  '2 : operator gedung
  '3 : UKP
  
  '---------
  
  Level1 = tbPengguna_getLevel1(frMenu1.nmLogin)
  Me.Text1 = ""
  Me.txtKarakkter = ","
  
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call dbMySQL_close
End Sub

Private Sub List1_DblClick()
  Dim pesan
  Dim namaFile As String, t1 As String
  Dim f
  Dim idx As Integer
  
  pesan = MsgBox("Simpan File Log ? ", vbYesNo)
  If pesan = vbYes Then
    Me.disable_Form
    namaFile = "d:\LogImportExcel-" & Year(Date) & "_" & Month(Date) & "_" & Day(Date) & " _ " & _
               "j" & Hour(Time) & Minute(Time) & Second(Time) & ".txt"
    Call OpenFile(namaFile, f, 2)
    For idx = 0 To List1.ListCount - 1
      List1.ListIndex = idx
      t1 = List1.text & Chr(13) & Chr(10)
      Call writefile(f, t1)
    Next
    Call closefile(f)
    MsgBox "File export di simpan di " & namaFile, vbInformation
    Me.Enable_Form
  End If
End Sub

