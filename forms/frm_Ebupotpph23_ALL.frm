VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Ebupotpph23_ALL 
   ClientHeight    =   7845
   ClientLeft      =   165
   ClientTop       =   210
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   12300
   Begin VB.ListBox List1 
      Height          =   480
      Left            =   120
      TabIndex        =   12
      Top             =   6960
      Width           =   11775
   End
   Begin VB.Frame Frame1 
      Caption         =   " Filter "
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   12015
      Begin VB.CommandButton cmd_sin 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Sinkronisasi"
         Height          =   375
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Sinkronisasi ke SPT dan SSP"
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cb_masa 
         Height          =   330
         Left            =   2760
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cb_tahun 
         Height          =   330
         Left            =   840
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox cb_NpwpKpp 
         Height          =   330
         Left            =   5160
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmd_load 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Export Template"
         Height          =   375
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Export ke Template ebupot. Siap di upload ke DJP"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "NPWP KPP"
         Height          =   210
         Left            =   4320
         TabIndex        =   8
         Top             =   420
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Masa"
         Height          =   210
         Left            =   2280
         TabIndex        =   7
         Top             =   420
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tahun "
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   12015
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4935
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   8705
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7590
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10557
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10557
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
      Caption         =   "Export eBupot 23"
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
Attribute VB_Name = "frm_Ebupotpph23_ALL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim nama_data As String
Dim LastNumberDasarPemotongan As Integer


Sub disable_Form()
    Me.Frame3.Enabled = False
    Me.Frame1.Enabled = False
End Sub

Sub Enable_Form()
    Me.Frame3.Enabled = True
    Me.Frame1.Enabled = True
End Sub


Sub format_Grid()
    
    Dim jRec As Long
    Dim c As Integer
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
        For c = 0 To rs.Fields.Count - 1
            'Me.DGrid1.Columns(c).Caption = UCase(rsGrid.Fields(c).Name)
            
            'kecil
            If c = 0 Or c = 1 Or c = 2 Then
                Me.DataGrid1.Columns(c).Alignment = dbgCenter
                Me.DataGrid1.Columns(c).Width = 400
            End If
            
            'If c = 12 Or c = 20 Then
            '    Me.DataGrid1.Columns(c).Alignment = dbgCenter
            '    Me.DataGrid1.Columns(c).NumberFormat = "dd mmm yy"
            '    Me.DataGrid1.Columns(c).Width = 900
            'End If
    
            If c = 7 Or c = 8 Then
                Me.DataGrid1.Columns(c).Alignment = dbgRight
                Me.DataGrid1.Columns(c).NumberFormat = "###,###"
                Me.DataGrid1.Columns(c).Width = 1400
            End If
        Next
End Sub



Private Sub cmd_export_Click()
    Dim jRec As Long
    
    Me.disable_Form
    jRec = RecordCount(rs)
    If jRec > 0 Then
        Call create_xls2(rs, "", "", "")
    End If
    Me.Enable_Form
End Sub


Private Sub LoadGrid(sql As String)
    Dim jRec As Long
    
    Me.disable_Form
    DoEvents
    
    If Trim(sql) <> "" Then
        Call dbMySQL_open
        If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
            sql = InputBox("", "", sql)
            MsgBox "error open sql"
            Me.Enable_Form
            Exit Sub
        End If
        'sql = InputBox("", "", sql)
        Set Me.DataGrid1.DataSource = rs
        jRec = RecordCount(rs)
    End If
    'Call format_Grid
    Me.Enable_Form
End Sub


Function load_ebupot23() As String
    Dim sql As String
    
    sql = "Select Masa_Pajak, Tahun_Pajak, Tanggal_Bukti_Potong, " & _
        "if(NPWP_WP_yang_Dipotong is null or NPWP_WP_yang_Dipotong= '' or left(NPWP_WP_yang_Dipotong,6) = '000000','N','Y') as bernpwp, " & _
        "NPWP_WP_yang_Dipotong, NIK_Yg_Dipotong, " & _
        "Nomer_telepon, Kode_Objek_Pajak, Penanda_tangan_BP_Pengurus, " & _
        "Jumlah_Nilai_Bruto_, Mendapatkan_Fasilitas, Nomor_SKB, " & _
        "Nomor_Aturan_DTP, NTPN_DTP, '23' as jenis_pajak, " & _
        "jenis_dokumen, No_Faktur_Pajak as no_dokumen, Tgl_Dokumen_ddMMyyyy as a2 " & _
        "From ebupot23 " & _
        "where NPWP_KPP = '" & Trim(Me.cb_npwpKpp.text) & _
        "' and Tahun_Pajak = '" & Trim(Me.cb_tahun.text) & _
        "' and Masa_Pajak = '" & Trim(Me.cb_masa.text) & "' " & _
        "order by No_Faktur_Pajak"

   load_ebupot23 = sql

End Function

Function load_ebupot23_sin() As String
    Dim sql As String
    
    sql = "Select " & _
        "NPWP_KPP, Kode_Form_Bukti_Potong, Masa_Pajak, " & _
        "Tahun_Pajak, Pembetulan, NPWP_WP_yang_Dipotong, " & _
        "Nama_WP_yang_dipotong, " & _
        "Alamat_WP_yang_Dipotong, Nomor_Bukti_Potong, Tanggal_Bukti_Potong, " & _
        "Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1, " & _
        "Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, " & _
        "Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, " & _
        "Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, " & _
        "Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, " & _
        "Nilai_Bruto_6a_Nilai_Bruto_6, Tarif_6a_Tarif_6, PPh_Yang_Dipotong__6a_PPh_Yang_Dipotong__6, " & _
        "Nilai_Bruto_6b_Nilai_Bruto_7, Tarif_6b_Tarif_7, PPh_Yang_Dipotong__6b_PPh_Yang_Dipotong__7, " & _
        "Nilai_Bruto_6c_Nilai_Bruto_8, Tarif_6c_Tarif_8, PPh_Yang_Dipotong__6c_PPh_Yang_Dipotong__8, " & _
        "Kode_Jasa_6d3_PMK_244_PMK03_2008, Nilai_Bruto_6d1, Tarif_6d1, " & _
        "PPh_Yang_Dipotong__6d1, Jumlah_Nilai_Bruto_, Jumlah_PPh_Yang_Dipotong, " & _
        "kode_divisi, Kode_Proyek, No_Bukti_Akuntansi, " & _
        "No_Faktur_Pajak, email " & _
        "From ebupot23 " & _
        "where NPWP_KPP = '" & Trim(Me.cb_npwpKpp.text) & _
        "' and Tahun_Pajak = '" & Trim(Me.cb_tahun.text) & _
        "' and Masa_Pajak = '" & Trim(Me.cb_masa.text) & "' " & _
        "order by No_Faktur_Pajak"

   load_ebupot23_sin = sql

End Function

Function load_ebupot26_sin() As String
    Dim sql As String
    
    sql = "Select " & _
        "NPWP_KPP, Kode_Form_Bukti_Potong, Masa_Pajak, " & _
        "Tahun_Pajak, Pembetulan, No_Paspor_WP_Terpotong, " & _
        "Nama_WP_yang_Dipotong, Alamat_WP_yang_Dipotong, Nomor_Bukti_Potong, " & _
        "Tanggal_Bukti_Potong,  " & _
        "Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1, " & _
        "Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, " & _
        "Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, " & _
        "Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, " & _
        "Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, " & _
        "Nilai_Bruto_6a_Nilai_Bruto_6, Tarif_6a_Tarif_6, PPh_Yang_Dipotong__6a_PPh_Yang_Dipotong__6, " & _
        "Nilai_Bruto_6b_Nilai_Bruto_7, Tarif_6b_Tarif_7, PPh_Yang_Dipotong__6b_PPh_Yang_Dipotong__7, " & _
        "Nilai_Bruto_6c_Nilai_Bruto_8, Tarif_6c_Tarif_8, PPh_Yang_Dipotong__6c_PPh_Yang_Dipotong__8, " & _
        "Kode_Jasa_6d1_PMK_244_PMK03_2008, Nilai_Bruto_6d1, Tarif_6d1, PPh_Yang_Dipotong__6d1, " & _
        "Jumlah_Nilai_Bruto_, Jumlah_PPh_Yang_Dipotong, kode_divisi, " & _
        "Kode_Proyek, No_Bukti_Akuntansi, No_Faktur_Pajak, " & _
        "email  " & _
        "From ebupot26 " & _
        "where NPWP_KPP = '" & Trim(Me.cb_npwpKpp.text) & _
        "' and Tahun_Pajak = '" & Trim(Me.cb_tahun.text) & _
        "' and Masa_Pajak = '" & Trim(Me.cb_masa.text) & "' " & _
        "order by No_Faktur_Pajak"

   load_ebupot26_sin = sql

End Function

Function load_ebupot26() As String
    Dim sql As String
    
    sql = "Select Masa_Pajak, Tahun_Pajak, Tanggal_Bukti_Potong, " & _
        "TIN_, Nama_WP_yang_Dipotong, tgl_lahir_wp, " & _
        "Alamat_WP_yang_Dipotong, No_Paspor_WP_Terpotong, No_Kitas_WP_Terpotong, " & _
        "Kode_Negara, Kode_Objek_Pajak, Penanda_tangan_BP_Pengurus, " & _
        "Jumlah_Nilai_Bruto_, Perkiraan_Penghasilan_Neto, Mendapatkan_Fasilitas, " & _
        "Nomor_Tanda_Terima_SKD, Tarif_SKD, Nomor_Aturan_DTP, " & _
        "NTPN_DTP, '26' as jenis_pajak, Jenis_Dokumen, " & _
        "No_Faktur_Pajak, Tgl_Dokumen_ as a2  " & _
        "From ebupot26 " & _
        "where NPWP_KPP = '" & Trim(Me.cb_npwpKpp.text) & _
        "' and Tahun_Pajak = '" & Trim(Me.cb_tahun.text) & _
        "' and Masa_Pajak = '" & Trim(Me.cb_masa.text) & "' " & _
        "order by No_Faktur_Pajak"
   load_ebupot26 = sql

End Function

Sub write_ebupot23(fileSimpan As String)
    Dim nmFile As String
    Dim f As String
    Dim fl As Object
    Dim fLs As Object
    Set fLs = CreateObject("Excel.Application")
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer
    
    nmFile = App.Path & "\rep\ebupot.xls"
    f = nmFile
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
        End If
    Else
        MsgBox "File template tidak ditemukan", vbCritical
    End If
    
    'open sheet 2, isi data looping
    Set fLs = fl.Sheets(2)
    baris = 3
    NO1 = 1
    
    If RecordCount(rs) > 0 Then
    
        rs.MoveFirst
        Do While rs.EOF = False
            kolom = 1
            fLs.Cells(baris, kolom).Value = NO1
            kolom = 2
            For a = 0 To 13
                If a = 0 Then
                    fLs.Cells(baris, kolom).NumberFormat = "@"
                    fLs.Cells(baris, kolom).Value = cek_null(rs(a))
                ElseIf a = 2 Then
                    fLs.Cells(baris, kolom).NumberFormat = "dd/MM/yyyy"
                    fLs.Cells(baris, kolom).HorizontalAlignment = Excel.xlRight
                    fLs.Cells(baris, kolom).Value = cek_null(rs(a))
                ElseIf a = 4 Or a = 5 Then
                    fLs.Cells(baris, kolom).NumberFormat = "@"
                    fLs.Cells(baris, kolom).Value = cek_null(rs(a))
                ElseIf a = 3 Or a = 8 Or a = 10 Then
                    fLs.Cells(baris, kolom).NumberFormat = "@"
                    fLs.Cells(baris, kolom).HorizontalAlignment = Excel.xlCenter
                    fLs.Cells(baris, kolom).Value = cek_null(rs(a))
                ElseIf a = 9 Then
                    fLs.Cells(baris, kolom).NumberFormat = "@"
                    fLs.Cells(baris, kolom).HorizontalAlignment = Excel.xlRight
                    fLs.Cells(baris, kolom).Value = cek_null(rs(a))
                Else
                    fLs.Cells(baris, kolom).Value = cek_null(rs(a))
                End If
                kolom = kolom + 1
            Next
            rs.MoveNext
            NO1 = NO1 + 1
            baris = baris + 1
        Loop
        
        '====dasar pemotongan - open sheet 4, isi data looping
        Set fLs = fl.Sheets(4)
        baris = 3
        NO1 = 1
        rs.MoveFirst
        Do While rs.EOF = False
            kolom = 1
            fLs.Cells(baris, kolom).Value = NO1
            fLs.Cells(baris, kolom).HorizontalAlignment = Excel.xlCenter
            kolom = 2
            For a = 14 To 17
                If a = 14 Then
                    'fLs.Cells(baris, kolom).NumberFormat = "General"
                    fLs.Cells(baris, kolom).Value = cek_null(rs(a))
                    fLs.Cells(baris, kolom).HorizontalAlignment = Excel.xlCenter
                ElseIf a = 15 Then
                    fLs.Cells(baris, kolom).NumberFormat = "@"
                    fLs.Cells(baris, kolom).HorizontalAlignment = Excel.xlCenter
                    fLs.Cells(baris, kolom).Value = cek_null(rs(a))
                ElseIf a = 16 Then
                    fLs.Cells(baris, kolom).NumberFormat = "@"
                    fLs.Cells(baris, kolom).Value = cek_null(rs(a))
                ElseIf a = 17 Then
                    fLs.Cells(baris, kolom).NumberFormat = "dd/MM/yyyy"
                    fLs.Cells(baris, kolom).HorizontalAlignment = Excel.xlCenter
                    fLs.Cells(baris, kolom).Value = cek_null(rs(a))
                Else
                    fLs.Cells(baris, kolom).Value = cek_null(rs(a))
                End If
                kolom = kolom + 1
            Next
            rs.MoveNext
            NO1 = NO1 + 1
            baris = baris + 1
        Loop
        
        LastNumberDasarPemotongan = NO1 - 1
        
        '-- rekap
        Set fLs = fl.Sheets(1)
        baris = 3
        kolom = 3
        fLs.Cells(baris, kolom).Value = LastNumberDasarPemotongan
    Else
        LastNumberDasarPemotongan = 0
        Me.List1.AddItem "ebupot23 : kosong"
        Me.List1.ListIndex = Me.List1.ListCount - 1
    End If
    
    'fl.ActiveWorkbook.Save
    
    fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
    'open by explorer
    'File1 = "explorer.exe " & fileSimpan
    'Call Shell(File1, vbNormalFocus)
    
End Sub


Sub write_ebupot26(fileSimpan As String)
    Dim nmFile As String
    Dim f As String
    Dim fl As Object
    Dim fLs As New Excel.Worksheet
    Dim baris As Integer, kolom As Integer, a As Integer
    Dim NO1 As Integer, jml26 As Integer
    
    nmFile = fileSimpan
    f = nmFile
    If is_file_ada(f) = True Then
        'File Valid
        If open_xls_lateBinding(fl, f) <> 0 Then
            Call pesan2("error open EXCEL", , vbYellow)
            Exit Sub
        End If
    Else
        MsgBox "File isian sebelumnya tidak ditemukan", vbCritical
        Exit Sub
    End If
    
    'open sheet 2, isi data looping
    Set fLs = fl.Sheets(3)
    baris = 3
    NO1 = 1
    rs.MoveFirst
    Do While rs.EOF = False
        kolom = 1
        fLs.Cells(baris, kolom).Value = NO1
        kolom = 2
        For a = 0 To 18
            If a = 0 Or a = 1 Then
                fLs.Cells(baris, kolom).Value = cek_null(rs(a))
            ElseIf a = 2 Or a = 5 Then
                fLs.Cells(baris, kolom).NumberFormat = "dd/MM/yyyy"
                fLs.Cells(baris, kolom).Value = cek_null(rs(a))
            Else
                fLs.Cells(baris, kolom).NumberFormat = "@"
                fLs.Cells(baris, kolom).Value = cek_null(rs(a))
            End If
            kolom = kolom + 1
        Next
        rs.MoveNext
        NO1 = NO1 + 1
        baris = baris + 1
    Loop
    
    '====dasar pemotongan - open sheet 4, isi data looping
    Set fLs = fl.Sheets(4)
    baris = LastNumberDasarPemotongan + 3
    NO1 = LastNumberDasarPemotongan + 1
    rs.MoveFirst
    jml26 = 0
    Do While rs.EOF = False
        jml26 = jml26 + 1
        kolom = 1
        fLs.Cells(baris, kolom).Value = NO1
        fLs.Cells(baris, kolom).HorizontalAlignment = Excel.xlCenter
        kolom = 2
        For a = 19 To 22
            If a = 14 Then
                'fLs.Cells(baris, kolom).NumberFormat = "General"
                fLs.Cells(baris, kolom).Value = cek_null(rs(a))
                fLs.Cells(baris, kolom).HorizontalAlignment = Excel.xlCenter
            ElseIf a = 15 Then
                fLs.Cells(baris, kolom).NumberFormat = "@"
                fLs.Cells(baris, kolom).HorizontalAlignment = Excel.xlCenter
                fLs.Cells(baris, kolom).Value = cek_null(rs(a))
            ElseIf a = 16 Then
                fLs.Cells(baris, kolom).NumberFormat = "@"
                fLs.Cells(baris, kolom).Value = cek_null(rs(a))
            ElseIf a = 17 Then
                fLs.Cells(baris, kolom).NumberFormat = "dd/MM/yyyy"
                fLs.Cells(baris, kolom).HorizontalAlignment = Excel.xlCenter
                fLs.Cells(baris, kolom).Value = cek_null(rs(a))
            Else
                fLs.Cells(baris, kolom).Value = cek_null(rs(a))
            End If
            kolom = kolom + 1
        Next
        rs.MoveNext
        NO1 = NO1 + 1
        baris = baris + 1
    Loop
    
    
    
    '-- rekap
    Set fLs = fl.Sheets(1)
    baris = 4
    kolom = 3
    fLs.Cells(baris, kolom).Value = jml26
    
    'fl.ActiveWorkbook.Save
    fl.ActiveWorkbook.SaveAs fileSimpan
    fl.Quit
    On Error Resume Next
    Call close_xls_lateBinding(fl)
    
    'open by explorer
    'File1 = "explorer.exe " & fileSimpan
    'Call Shell(File1, vbNormalFocus)
End Sub

Private Sub cmd_load_Click()
    Dim sql As String
    Dim fileSimpan As String, File1 As String
    
    On Error GoTo er1
    MsgBox "Pada saat proses selesai, kadang kala file excel tidak mau tampil. " & vbCr & _
            "Lakukan alt+tab. Jika ada konfirmasi, pilih 'switch'", vbInformation
    Me.disable_Form
    
    fileSimpan = App.Path & "\exp\" & Trim(Me.cb_npwpKpp.text) & ".xls"
    
    Me.List1.Clear
    Me.List1.AddItem "File simpan:" & fileSimpan
    Me.List1.ListIndex = Me.List1.ListCount - 1
    Me.List1.AddItem "Load ebupot23"
    Me.List1.ListIndex = Me.List1.ListCount - 1
    'load query ebupot23
    sql = load_ebupot23
    Call LoadGrid(sql)
    
    'masukkan ke xls template
    Me.List1.AddItem "write ebupot23"
    Me.List1.ListIndex = Me.List1.ListCount - 1
    Call write_ebupot23(fileSimpan)
    Call delay(2)
    
    'load query ebupot26
    Me.List1.AddItem "Load ebupot26"
    Me.List1.ListIndex = Me.List1.ListCount - 1
    sql = load_ebupot26
    Call LoadGrid(sql)
    Call delay(2)
    
    If RecordCount(rs) > 0 Then
        Me.List1.AddItem "write ebupot26"
        Me.List1.ListIndex = Me.List1.ListCount - 1
        Call write_ebupot26(fileSimpan)
    Else
        Me.List1.AddItem "ebupot26 - no data"
        Me.List1.ListIndex = Me.List1.ListCount - 1
    End If
    'masukkan ke xls
    'tulis rekap
    
    'open file
    'open by explorer
    File1 = "explorer.exe " & fileSimpan
    Call Shell(File1, vbNormalFocus)
    
    'done
    Me.Enable_Form
    MsgBox "Proses export selesai. " & vbCr & "File di :" & fileSimpan, vbInformation
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
End Sub


Sub sin_ebupot23()
Dim sql As String
    Dim rs As ADODB.Recordset, jRec As Long, c As Long
    
    Dim npwp_kpp As String, Kode_Form As String, Masa_Pajak As String, Tahun_Pajak As String
    Dim Pembetulan As String, npwp_wp As String, Nama_WP As String, Alamat_WP As String
    Dim Nomor_Bukti_Potong As String, Tanggal_Bukti_Potong As Date
    Dim Nilai_Bruto_1 As Currency, Tarif_1 As String, PPh_Yang_Dipotong__1 As Currency
    Dim Nilai_Bruto_2 As Currency, Tarif_2 As String, PPh_Yang_Dipotong__2 As Currency
    Dim Nilai_Bruto_3 As Currency, Tarif_3 As String, PPh_Yang_Dipotong__3 As Currency
    Dim Nilai_Bruto_4 As Currency, Tarif_4 As String, PPh_Yang_Dipotong__4 As Currency
    Dim Nilai_Bruto_5 As Currency, Tarif_5 As String, PPh_Yang_Dipotong__5 As Currency
    Dim Nilai_Bruto_6a As Currency, Tarif_6a As String, PPh_Yang_Dipotong__6a As Currency
    Dim Nilai_Bruto_6b As Currency, Tarif_6b As String, PPh_Yang_Dipotong__6b As Currency
    Dim Nilai_Bruto_6c As Currency, Tarif_6c As String, PPh_Yang_Dipotong__6c As Currency
    Dim Kode_Jasa_6d1 As String, Nilai_Bruto_6d1 As Currency, Tarif_6d1 As String, PPh_Yang_Dipotong__6d1 As Currency
    Dim Jumlah_Nilai_Bruto_ As Currency, Jumlah_PPh_Yang_Dipotong As Currency, kode_divisi As String
    Dim kd_proyek As String, nott As String, nofaktur As String, email As String
    
    Dim data_Valid As Boolean
    Dim jml_Insert As Long, jml_Update As Long
    Dim return1 As Integer, mod1 As Integer
    
    sql = load_ebupot23_sin
    
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rs)
    If jRec > 0 Then
        rs.MoveFirst
        c = 1
        Do While rs.EOF = False
            Call info_progress(Me.StatusBar1, 1, c, jRec, "sinkronisasi ebupot23")
            
            mod1 = c Mod 1000
            If mod1 = 0 Then Call dbMySQL_open
            
            npwp_kpp = cek_null(rs(0))
            Kode_Form = cek_null(rs(1))
            Masa_Pajak = cek_null(rs(2))
            Tahun_Pajak = cek_null(rs(3))
            Pembetulan = cek_null(rs(4))
            npwp_wp = cek_null(rs(5))
            Nama_WP = cek_null(rs(6))
            Alamat_WP = cek_null(rs(7))
            Nomor_Bukti_Potong = cek_null(rs(8))
            Tanggal_Bukti_Potong = cek_Date(rs(9))
            Nilai_Bruto_1 = cek_null(rs(10))
            Tarif_1 = cek_null(rs(11))
            PPh_Yang_Dipotong__1 = cek_null(rs(12))
            Nilai_Bruto_2 = cek_null(rs(13))
            Tarif_2 = cek_null(rs(14))
            PPh_Yang_Dipotong__2 = cek_null(rs(15))
            Nilai_Bruto_3 = cek_null(rs(16))
            Tarif_3 = cek_null(rs(17))
            PPh_Yang_Dipotong__3 = cek_null(rs(18))
            Nilai_Bruto_4 = cek_null(rs(19))
            Tarif_4 = cek_null(rs(20))
            PPh_Yang_Dipotong__4 = cek_null(rs(21))
            Nilai_Bruto_5 = cek_null(rs(22))
            Tarif_5 = cek_null(rs(23))
            PPh_Yang_Dipotong__5 = cek_null(rs(24))
            Nilai_Bruto_6a = cek_null(rs(25))
            Tarif_6a = cek_null(rs(26))
            PPh_Yang_Dipotong__6a = cek_null(rs(27))
            Nilai_Bruto_6b = cek_null(rs(28))
            Tarif_6b = cek_null(rs(29))
            PPh_Yang_Dipotong__6b = cek_null(rs(30))
            Nilai_Bruto_6c = cek_null(rs(31))
            Tarif_6c = cek_null(rs(32))
            PPh_Yang_Dipotong__6c = cek_null(rs(33))
            Kode_Jasa_6d1 = cek_null(rs(34))
            Nilai_Bruto_6d1 = cek_null(rs(35))
            Tarif_6d1 = cek_null(rs(36))
            PPh_Yang_Dipotong__6d1 = cek_null(rs(37))
            Jumlah_Nilai_Bruto_ = cek_null(rs(38))
            Jumlah_PPh_Yang_Dipotong = cek_null(rs(39))
            kode_divisi = cek_null(rs(40))
            kd_proyek = cek_null(rs(41))
            nott = cek_null(rs(42))
            nofaktur = cek_null(rs(43))
            email = cek_null(rs(44))
            
            data_Valid = True
            If Trim(Nomor_Bukti_Potong) <> "" Then
                '-----------
                If CStr(Year(Tanggal_Bukti_Potong)) = Tahun_Pajak Then
                Else
                    data_Valid = False
                    Call setListInfo(Me.List1, "Data ke " & c & " Tahun Pajak tidak valid")
                End If
                
                If adddigit(CLng(Month(Tanggal_Bukti_Potong)), 2) = adddigit(cek_Lng(Masa_Pajak), 2) Then
                Else
                    data_Valid = False
                    Call setListInfo(Me.List1, "Data ke " & c & " Masa Pajak tidak valid")
                End If
                '-----------
                
                If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
                    data_Valid = False
                    Me.List1.AddItem "Data ke " & c & " NPWP_KPP tidak valid"
                ElseIf Jumlah_PPh_Yang_Dipotong = 0 Then
                    data_Valid = False
                    Me.List1.AddItem "Data ke " & c & " Nilai PPH 0"
                End If
                
                If data_Valid = True Then
                
                    return1 = tbPph23_insert(npwp_kpp, Kode_Form, Masa_Pajak, Tahun_Pajak, Pembetulan, npwp_wp, Nama_WP, _
                                        Alamat_WP, Nomor_Bukti_Potong, Tanggal_Bukti_Potong, _
                                        Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1, _
                                        Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, _
                                        Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, _
                                        Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, _
                                        Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, _
                                        Nilai_Bruto_6a, Tarif_6a, PPh_Yang_Dipotong__6a, _
                                        Nilai_Bruto_6b, Tarif_6b, PPh_Yang_Dipotong__6b, _
                                        Nilai_Bruto_6c, Tarif_6c, PPh_Yang_Dipotong__6c, _
                                        Kode_Jasa_6d1, Nilai_Bruto_6d1, Tarif_6d1, PPh_Yang_Dipotong__6d1, _
                                        Jumlah_Nilai_Bruto_, Jumlah_PPh_Yang_Dipotong, kode_divisi, kd_proyek, nott, nofaktur, email)
                    If return1 = 1 Then
                        jml_Insert = jml_Insert + 1
                    ElseIf return1 = 2 Then
                        jml_Update = jml_Update + 1
                    Else
                        Call pesan2("Insert error", , vbYellow)
                        Exit Sub
                    End If
                End If
            End If
            
            rs.MoveNext
            c = c + 1
        Loop
        
        MsgBox "ebupot23 total data:" & jRec & vbCr & _
                "total insert:" & jml_Insert & _
                "total update:" & jml_Update, vbInformation
        
    Else
        Me.List1.AddItem "Sinkronisasi ebupot23 - jml Data :0"
        Me.List1.ListIndex = Me.List1.ListCount - 1
    End If
End Sub

Sub sin_ebupot26()
    Dim sql As String
    Dim rs As ADODB.Recordset, jRec As Long, c As Long
    
    Dim npwp_kpp As String, Kode_Form As String, Masa_Pajak As String, Tahun_Pajak As String, _
                    Pembetulan As String, npwp_wp As String, Nama_WP As String, Alamat_WP As String, _
                    Nomor_Bukti_Potong As String, Tanggal_Bukti_Potong As Date, _
                    Nilai_Bruto_1 As Currency, Tarif_1 As String, PPh_Yang_Dipotong__1 As Currency, _
                    Nilai_Bruto_2 As Currency, Tarif_2 As String, PPh_Yang_Dipotong__2 As Currency, _
                    Nilai_Bruto_3 As Currency, Tarif_3 As String, PPh_Yang_Dipotong__3 As Currency, _
                    Nilai_Bruto_4 As Currency, Tarif_4 As String, PPh_Yang_Dipotong__4 As Currency, _
                    Nilai_Bruto_5 As Currency, Tarif_5 As String, PPh_Yang_Dipotong__5 As Currency, _
                    Nilai_Bruto_6a As Currency, Tarif_6a As String, PPh_Yang_Dipotong__6a As Currency, _
                    Nilai_Bruto_6b As Currency, Tarif_6b As String, PPh_Yang_Dipotong__6b As Currency, _
                    Nilai_Bruto_6c As Currency, Tarif_6c As String, PPh_Yang_Dipotong__6c As Currency, _
                    Kode_Jasa_6d1 As String, Nilai_Bruto_6d1 As Currency, Tarif_6d1 As String, PPh_Yang_Dipotong__6d1 As Currency, _
                    Jumlah_Nilai_Bruto_ As Currency, Jumlah_PPh_Yang_Dipotong As Currency, kode_divisi As String, _
                    kd_proyek As String, nott As String, nofaktur As String, email As String
    
    Dim data_Valid As Boolean
    Dim jml_Insert As Long, jml_Update As Long
    Dim return1 As Integer, mod1 As Integer
    
    sql = load_ebupot26_sin
    
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rs)
    If jRec > 0 Then
        rs.MoveFirst
        c = 1
        Do While rs.EOF = False
            Call info_progress(Me.StatusBar1, 1, c, jRec, "sinkronisasi ebupot23")
            
            mod1 = c Mod 1000
            If mod1 = 0 Then Call dbMySQL_open
            
            npwp_kpp = cek_null(rs(0))
            Kode_Form = cek_null(rs(1))
            Masa_Pajak = cek_null(rs(2))
            Tahun_Pajak = cek_null(rs(3))
            Pembetulan = cek_null(rs(4))
            npwp_wp = cek_null(rs(5))
            If Trim(npwp_wp) = "" Then
                npwp_wp = "000000000000000"
            End If
            Nama_WP = cek_null(rs(6))
            Alamat_WP = cek_null(rs(7))
            Nomor_Bukti_Potong = cek_null(rs(8))
            Tanggal_Bukti_Potong = cek_null(rs(9))
            Nilai_Bruto_1 = cek_null(rs(10))
            Tarif_1 = cek_null(rs(11))
            PPh_Yang_Dipotong__1 = cek_null(rs(12))
            Nilai_Bruto_2 = cek_null(rs(13))
            Tarif_2 = cek_null(rs(14))
            PPh_Yang_Dipotong__2 = cek_null(rs(15))
            Nilai_Bruto_3 = cek_null(rs(16))
            Tarif_3 = cek_null(rs(17))
            PPh_Yang_Dipotong__3 = cek_null(rs(18))
            Nilai_Bruto_4 = cek_null(rs(19))
            Tarif_4 = cek_null(rs(20))
            PPh_Yang_Dipotong__4 = cek_null(rs(21))
            Nilai_Bruto_5 = cek_null(rs(22))
            Tarif_5 = cek_null(rs(23))
            PPh_Yang_Dipotong__5 = cek_null(rs(24))
            Nilai_Bruto_6a = cek_null(rs(25))
            Tarif_6a = cek_null(rs(26))
            PPh_Yang_Dipotong__6a = cek_null(rs(27))
            Nilai_Bruto_6b = cek_null(rs(28))
            Tarif_6b = cek_null(rs(29))
            PPh_Yang_Dipotong__6b = cek_null(rs(30))
            Nilai_Bruto_6c = cek_null(rs(31))
            Tarif_6c = cek_null(rs(32))
            PPh_Yang_Dipotong__6c = cek_null(rs(33))
            Kode_Jasa_6d1 = cek_null(rs(34))
            Nilai_Bruto_6d1 = cek_null(rs(35))
            Tarif_6d1 = cek_null(rs(36))
            PPh_Yang_Dipotong__6d1 = cek_null(rs(37))
            Jumlah_Nilai_Bruto_ = cek_null(rs(38))
            Jumlah_PPh_Yang_Dipotong = cek_null(rs(39))
            kode_divisi = cek_null(rs(40))
            kd_proyek = cek_null(rs(41))
            nott = cek_null(rs(42))
            nofaktur = cek_null(rs(43))
            email = cek_null(rs(44))
            
            data_Valid = True
            If Trim(Nomor_Bukti_Potong) <> "" Then
                '-----------
                If CStr(Year(Tanggal_Bukti_Potong)) = Tahun_Pajak Then
                Else
                    data_Valid = False
                    Call setListInfo(Me.List1, "Data ke " & c & " Tahun Pajak tidak valid")
                End If
                
                If adddigit(CLng(Month(Tanggal_Bukti_Potong)), 2) = adddigit(cek_Lng(Masa_Pajak), 2) Then
                Else
                    data_Valid = False
                    Call setListInfo(Me.List1, "Data ke " & c & " Masa Pajak tidak valid")
                End If
                '-----------
                
                If tbMKpp_isNpwpKPP_Valid(npwp_kpp) = False Then
                    data_Valid = False
                    Me.List1.AddItem "Data ke " & c & " NPWP_KPP tidak valid"
                ElseIf Jumlah_PPh_Yang_Dipotong = 0 Then
                    data_Valid = False
                    Me.List1.AddItem "Data ke " & c & " Nilai PPH 0"
                End If
                
                If data_Valid = True Then
                
                    return1 = tbPph26_insert(npwp_kpp, Kode_Form, Masa_Pajak, Tahun_Pajak, Pembetulan, _
                                        npwp_wp, Nama_WP, _
                                        Alamat_WP, Nomor_Bukti_Potong, Tanggal_Bukti_Potong, _
                                        Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1, _
                                        Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, _
                                        Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, _
                                        Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, _
                                        Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, _
                                        Nilai_Bruto_6a, Tarif_6a, PPh_Yang_Dipotong__6a, _
                                        Nilai_Bruto_6b, Tarif_6b, PPh_Yang_Dipotong__6b, _
                                        Nilai_Bruto_6c, Tarif_6c, PPh_Yang_Dipotong__6c, _
                                        Kode_Jasa_6d1, Nilai_Bruto_6d1, Tarif_6d1, PPh_Yang_Dipotong__6d1, _
                                        Jumlah_Nilai_Bruto_, Jumlah_PPh_Yang_Dipotong, kode_divisi, _
                                        kd_proyek, nott, nofaktur, email)
                    If return1 = 1 Then
                        jml_Insert = jml_Insert + 1
                    ElseIf return1 = 2 Then
                        jml_Update = jml_Update + 1
                    Else
                        Call pesan2("Insert error", , vbYellow)
                        Exit Sub
                    End If
                End If
            End If
            
            rs.MoveNext
            c = c + 1
        Loop
        
        MsgBox "ebupot26 total data:" & jRec & vbCr & _
                "total insert:" & jml_Insert & _
                "total update:" & jml_Update, vbInformation
        
    Else
        Me.List1.AddItem "Sinkronisasi ebupot26 - jml Data :0"
        Me.List1.ListIndex = Me.List1.ListCount - 1
    End If
End Sub

Private Sub cmd_sin_Click()
    
    
    Me.List1.Clear
    

    Me.List1.AddItem "Sinkronisasi ebupot23"
    Me.List1.ListIndex = Me.List1.ListCount - 1
    Call sin_ebupot23
    Me.List1.AddItem "Sinkronisasi ebupot26"
    Me.List1.ListIndex = Me.List1.ListCount - 1
    Call sin_ebupot26

    
    
End Sub

Private Sub Form_Load()
  Dim sql As String
  Dim Level1 As Integer
  
  nama_data = "ebupot23"
  Call dbMySQL_open
    
  'load combo
  
  sql = "select distinct Tahun_Pajak from " & _
        "( select Tahun_Pajak from ebupot23 Union " & _
        "select Tahun_Pajak from ebupot26)as a"
  Call Load_combo(Me.cb_tahun, sql, cnn, True, , 0)
  
  sql = "select distinct Masa_Pajak from " & _
        "(select Masa_Pajak from ebupot23 Union  " & _
        "select Masa_Pajak from ebupot26 )as a"
  Call Load_combo(Me.cb_masa, sql, cnn, True, , 0)
  
  sql = "select distinct NPWP_KPP from " & _
        "( select NPWP_KPP from ebupot23 Union  " & _
        "select NPWP_KPP from ebupot26 )as a"
  Call Load_combo(Me.cb_npwpKpp, sql, cnn, True, , 0)
  
  Me.Height = 8310
  Me.Width = 12420
  
  'get level
  '2 : operator gedung
  '3 : UKP
  
  '---------
  
  'Level1 = tbPengguna_getLevel1(frMenu1.nmLogin)
  'If Level1 = 2 Then
  '  Me.cb_divisi.text = tbPengguna_getDivisi(frMenu1.nmLogin)
  '  Me.cb_divisi.Enabled = False
  'ElseIf Level1 = 3 Then
  'Else
  '  Call pesan2("Level tidak valid", , vbYellow)
  ' Me.cb_divisi.Enabled = False
  'End If
 
 'Call LoadGrid
  Call pesan2("Pilih Filter dan klik 'Proses', atau " & vbCr & _
                "klik cari data dan ENTER")
End Sub


Private Sub Form_Resize()
    If Me.Width - 405 > 0 Then Me.Frame3.Width = Me.Width - 405
    If Me.Height - 2595 > 0 Then Me.Frame3.Height = Me.Height - 2895

    If Me.Width - 645 > 0 Then Me.DataGrid1.Width = Me.Width - 645
    If Me.Height - 3435 > 0 Then Me.DataGrid1.Height = Me.Height - 3375

    If Me.Height - 1050 > 0 Then Me.List1.Top = Me.Height - 1350
    'If Me.Width - 12300 > 0 Then Me.List1.Left = Me.Width - 12300

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call dbMySQL_close
End Sub


