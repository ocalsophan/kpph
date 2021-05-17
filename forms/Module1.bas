Attribute VB_Name = "Module1"
Option Explicit
Public cnn As New ADODB.connection
Public cnnTemp As New ADODB.connection

Public Type SHFILEOPSTRUCT
     hWnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Integer
     fAnyOperationsAborted As Boolean
     hNameMappings As Long
     lpszProgressTitle As String
End Type

Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Const FO_COPY = &H2
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_ALLOWUNDO = &H40


'------------ variabel public LOV
Public lov_return As TextBox            'menyimpan nilai kembalian LOV
Public lov_SqL As String                'isi SQL: select kd_grup, nm_group form gvendor
Public lov_kolom_Dicari As String       'pisah dgn koma : kd_grup, nm_grup
Public lov_order_by As String           'lov_order_by: nama kolom order by
Public lov_Key_Cari As String
Public lov_title As String
Public lov_width_field1 As Integer
Public lov_width_field2 As Integer
'-----------

Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long

Public Function isddMMyyyy(txt As String) As Boolean
    'cek apakah format sudah mengikuti dd/MM/yyyy ?
    
    'cek panjang
    'cek simbol baca
    'cek tgl antara 1 s/d 31
    'cek bulan antara 1 s/d 12
    'cek tahun antara 1990 s/d tahun sekarang
    Dim tgl As Integer, bulan As Integer, tahun As Integer
    
    txt = Trim(txt)
    If Len(txt) <> 10 Then
        isddMMyyyy = False
        Exit Function
    ElseIf (Mid(txt, 3, 1) <> "/" Or Mid(txt, 6, 1) <> "/") Then
        isddMMyyyy = False
        Exit Function
    Else
        tgl = cek_Int(Left(txt, 2))
        bulan = cek_Int(Mid(txt, 4, 2))
        tahun = cek_Int(Right(txt, 4))
        
        If tgl < 1 Or tgl > 31 Then
            
        ElseIf bulan < 1 Or bulan > 12 Then
            isddMMyyyy = False
            Exit Function
        ElseIf tahun < 1990 Or tahun > Year(Now) Then
            isddMMyyyy = False
            Exit Function
        Else
            isddMMyyyy = True
            Exit Function
        End If
    End If
End Function



Public Function IsOnline() As Boolean
  Dim LFlags As Long
  IsOnline = InternetGetConnectedState(LFlags, 0&)
End Function

Sub setListInfo(ByRef List1 As ListBox, pesan As String)
    List1.AddItem pesan
    List1.ListIndex = List1.ListCount - 1
End Sub
            
Function get_data(ByRef cnn1 As ADODB.connection, namaTabel As String, kolom As String, where1 As String, Optional isResAngka As Boolean = False) As String
    Dim sql As String, t As String
    
    sql = "select " & kolom & " from " & namaTabel & " where " & where1
    t = cari_data1(cnn1, sql, isResAngka)
    get_data = t
End Function

Function del_Data(ByRef cnn1 As ADODB.connection, namaTabel As String, where1 As String) As Boolean
    Dim sql As String, t As String
    
    sql = "delete from " & namaTabel & " where " & where1
    If ExecSQL1(cnn, sql) <> 0 Then
        del_Data = False
    Else
        del_Data = True
    End If
End Function

Function Is_DirExists(OrigFile As String) As Boolean
  'Returns a boolean - True if the folder exists
  Dim fs
  Set fs = CreateObject("Scripting.FileSystemObject")
  Is_DirExists = fs.folderexists(OrigFile)
End Function


Function isDataAda2(ByRef c As ADODB.connection, nmTabel As String, kondisi As String) As Boolean
    'return true jika data ada
    Dim sql As String, t As String
    
    If Trim(kondisi) = "" Then
        isDataAda2 = False
    Else
        sql = "select count(*) from " & Trim(nmTabel) & " where " & kondisi
        t = cari_data1(c, sql, True)
        If CInt(t) > 0 Then
            isDataAda2 = True
        Else
            isDataAda2 = False
        End If
    End If
    
        
End Function

Function cek_format_tanggal() As Boolean
    Dim benar As Boolean
    Dim t As String, t2 As String
    
    benar = True
    t = CStr(Date)
    t2 = Format(Date, "dd/mm/yyyy")
    If Trim(t) <> Trim(t2) Then
        benar = False
    End If
    
    If benar = True Then
        t2 = Format(CDate("2015-10-25"), "dd/mmm/yyyy")
        If t2 <> "25/Okt/2015" Then
            benar = False
        End If
    End If
   cek_format_tanggal = benar
End Function

Function getTimeStamp(Optional dt As Date) As String
    Dim tgl As String, bln As String, thn As String
    Dim jam As String, menit As String, detik As String
    Dim t As String
    
    tgl = Trim(str(Day(dt)))
    tgl = adddigit(CLng(tgl), 2)
  
    bln = Trim(str(Month(dt)))
    bln = adddigit(CLng(bln), 2)
    
    thn = Right(Trim(str(Year(dt))), 2)
    jam = Trim(str(Hour(Time)))
    jam = adddigit(CLng(jam), 2)
    menit = Trim(str(Minute(Time)))
    menit = adddigit(CLng(menit), 2)
    detik = Trim(Second(Time))
    detik = adddigit(CLng(detik), 2)
  
    t = thn & bln & tgl & "_" & jam & menit & detik
    getTimeStamp = t
End Function

Function get_versiapp() As String
    Dim t As String
    t = App.Major & App.Minor & App.Revision
    get_versiapp = t
End Function



Function adddigit(angka As Long, jml_digit As Integer) As String
  Dim t As String
  t = Trim(str(angka))
  Do While Len(Trim(t)) < jml_digit
    t = "0" + Trim(t)
  Loop
  adddigit = t
End Function

Sub simpanIsiListBox(List1 As ListBox)
  Dim pesan
  Dim namaFile As String, t1 As String
  Dim f
  Dim idx As Integer
  
  pesan = MsgBox("Simpan File Log ? ", vbYesNo)
  If pesan = vbYes Then
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
  End If
End Sub


Function cek_null(a, Optional pengganti = "") As String
  On Error GoTo er1
  
  
    If IsNull(a) = False Then
        If Trim(a) <> "" Then
            cek_null = Trim(a)
        Else
            cek_null = pengganti
        End If
    Else
        cek_null = pengganti
    End If
  Exit Function
er1:
  cek_null = "[error]"
End Function

Function cek_Date(a) As Date
  Dim pengganti As Date
  On Error GoTo er1
    
    pengganti = DateValue(Now)
    If IsNull(a) = False Then
        If isDate(a) = True Then
            cek_Date = a
        Else
            cek_Date = pengganti
        End If
    Else
        cek_Date = pengganti
    End If
  Exit Function
er1:
  cek_Date = pengganti
End Function


Function cek_Int(a, Optional t As Integer = 0) As Integer
  On Error GoTo er1
  If IsNumeric(a) = True Then
    cek_Int = CInt(a)
  Else
    cek_Int = t
  End If
  Exit Function
er1:
    cek_Int = t
End Function

Function cek_Lng(a, Optional t As Long = 0) As Long
  On Error GoTo er1
  If IsNumeric(a) = True Then
    cek_Lng = CLng(a)
  Else
    cek_Lng = t
  End If
  Exit Function
er1:
    cek_Lng = t
End Function

Function cek_Dbl(a) As String
  On Error GoTo er1
  If IsNumeric(a) = True Then
    cek_Dbl = Replace(a, ",", ".")
  Else
    cek_Dbl = "0"
  End If
  Exit Function
er1:
    cek_Dbl = 0
End Function

Function cek_Money(a) As Currency
  On Error GoTo er1
  If IsNumeric(a) = True Then
    cek_Money = CCur(a)
  Else
    cek_Money = 0
  End If
  Exit Function
er1:
    cek_Money = 0
End Function

Function cek_Currency(a, Optional pengganti As Currency = 0) As Currency
  On Error GoTo er1
    
    If IsNull(a) = False Then
        If IsNumeric(a) = True Then
            cek_Currency = a
        Else
            cek_Currency = pengganti
        End If
    Else
        cek_Currency = pengganti
    End If
  Exit Function
er1:
  cek_Currency = pengganti
End Function


Function set_tgl_perv(dt As Date) As String
  Dim tgl, bulan, thn As String
  
  tgl = Day(dt)
  thn = Year(dt)
  bulan = Month(dt)
  
  Do While Len(RTrim(tgl)) < 2
    tgl = "0" & tgl
  Loop
  Do While Len(RTrim(bulan)) < 2
    bulan = "0" & bulan
  Loop

  set_tgl_perv = thn & "-" & bulan & "-" & tgl
End Function


Public Sub show_LOV2(sql As String, kolom_dicari As String, order_by As String, title As String, keyReturn As TextBox, _
       keyCari As String, lbr_field1 As Integer, lbr_field2 As Integer, Optional left1 As Integer = -1, Optional top1 As Integer = -1)
  
  Set lov_return = keyReturn
  
  lov_SqL = sql
  lov_kolom_Dicari = kolom_dicari
  lov_order_by = order_by
  lov_title = title
  lov_Key_Cari = keyCari
  lov_width_field1 = lbr_field1
  lov_width_field2 = lbr_field2
  LOV_2.Show vbModal
  'If left1 > 0 And top1 > 0 Then
  '  LOV_2.Left = left1
  '  LOV_2.Top = top1
  'End If
End Sub


Function load_setup(ByRef dsn_Ver As String, _
            ByRef dsn_SPM As String) As Integer

  'load info nama server utk database verifikasi dan cashbasis
  
  Dim msg As String
  Dim MyVar
  Dim f
  Dim fso
    
  On Error GoTo er1
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Set fso = CreateObject("Scripting.FileSystemObject")
  If (fso.FileExists(App.Path & "\data\setup.txt")) Then
    Set f = fso.OpenTextFile(App.Path & "\data\setup.txt", ForReading, True)
    load_setup = 0
  Else
    load_setup = 1
    MsgBox ("File Setup tidak ada")
    GoTo er1
  End If
  
  'nama server
  dsn_Ver = f.readline
  If dsn_Ver = "" Then GoTo er1
  dsn_SPM = f.readline
  If dsn_SPM = "" Then GoTo er1
  
  f.Close
  Exit Function
er1:
  load_setup = 1
End Function

Public Function hapusPetik(txtInput As String) As String
    txtInput = Replace(txtInput, "'", "")
    hapusPetik = txtInput
End Function

Public Function CekPetik(TextInput As String) As String
  Dim h As Integer
  Dim voutput As String

  voutput = Trim(TextInput)
  h = 1
  While h <> 0
    h = InStr(h, voutput, "'", 1)
    If h > 0 Then
      voutput = Left(voutput, h) & "'" & Right(voutput, Len(voutput) - h)
      h = h + 2
    End If
  
  Wend
  CekPetik = voutput
End Function

Function isDataAda(nmTabel As String, nmKolomWhere As String, strDicari As String, cnn As ADODB.connection, _
                    Optional isDate As Boolean = False) As Boolean
    'return true jika data ada
    Dim sql As String, t As String
    
    If isDate = False Then
    
        sql = "select count(*) from " & Trim(nmTabel) & " where ucase(" & Trim(nmKolomWhere) & _
                ") = '" & Trim(UCase(strDicari)) & "'"
    Else
        sql = "select count(*) from " & Trim(nmTabel) & " where " & Trim(nmKolomWhere) & _
                " = '" & Trim(UCase(strDicari)) & "'"
    End If
    t = cari_data1(cnn, sql, True)
    If CInt(t) > 0 Then
        isDataAda = True
    Else
        isDataAda = False
    End If
End Function


Public Function get_kode_combo(cb As ComboBox, karakterbatas As String)
     Dim t As String, t2 As String
     Dim a As Integer
     
     t = cb.text
     'ambil text s/d karakter batas
     t2 = ""
     For a = 1 To Len(t)
          If Mid(t, a, 1) = Trim(karakterbatas) Then
               Exit For
          Else
               t2 = t2 & Mid(t, a, 1)
          End If
     Next
     get_kode_combo = Trim(t2)
End Function

Public Function GantiPetik(TextInput As String) As String
  Dim h As Integer
  Dim voutput As String

  voutput = Trim(TextInput)
  
  'petik diganti
  h = 1
  While h <> 0
    h = InStr(h, voutput, "'", 1)
    If h > 0 Then
      voutput = Left(voutput, h - 1) & " " & Right(voutput, Len(voutput) - h)
      h = h + 2
    End If
  Wend


  GantiPetik = voutput
End Function

Function InStrArray(valDiCari As String, arr()) As Boolean
   Dim ret1 As Boolean
   ret1 = False
 
   Dim i As Integer
   For i = LBound(arr) To UBound(arr)
        If arr(i) = valDiCari Then
            ret1 = True
            Exit For
        End If
   Next
   InStrArray = ret1
End Function

Public Function cekStringAngka(input1 As String) As String
    'baca dari kiri ke kanan
    Dim c As Integer
    Dim tBaru As String
    
    
    tBaru = ""
    For c = 1 To Len(input1)
        If IsNumeric(Mid(input1, c, 1)) = True Then
            tBaru = tBaru & Mid(input1, c, 1)
        End If
    Next
    cekStringAngka = tBaru
End Function


Public Function CekInputString(TextInput As String) As String
  Dim h As Integer
  Dim voutput As String

  voutput = Trim(TextInput)
  
  'petik ditambahi
  voutput = CekPetik(voutput)
  
  'koma diganti -
  h = 1
  While h <> 0
    h = InStr(h, voutput, ",", 1)
    If h > 0 Then
      voutput = Left(voutput, h - 1) & " - " & Right(voutput, Len(voutput) - h)
      h = h + 2
    End If
  Wend

  '; diganti -
  h = 1
  While h <> 0
    h = InStr(h, voutput, ";", 1)
    If h > 0 Then
      voutput = Left(voutput, h - 1) & " - " & Right(voutput, Len(voutput) - h)
      h = h + 2
    End If
  Wend

  CekInputString = voutput
End Function

Sub OpenFile(nmFile As String, ByRef f, mode)
  Dim msg As String
  Dim MyVar
  Dim fso
  
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.OpenTextFile(nmFile, mode, True)
End Sub

Sub writefile(ByRef f, txt As String)
  f.Write (txt)
End Sub

Sub closefile(ByRef f)
  On Error Resume Next
  f.Close
End Sub

Sub create_ds_Access(file_DSN As String, folder_mdb As String, file_mdb As String)
  Dim f
  
  OpenFile file_DSN, f, 2
  writefile f, "[ODBC]" & Chr(13) & Chr(10)
  writefile f, "DRIVER=Microsoft Access Driver (*.mdb)" & Chr(13) & Chr(10)
  writefile f, "UID = admin" & Chr(13) & Chr(10)
  writefile f, "UserCommitSync = Yes" & Chr(13) & Chr(10)
  writefile f, "Threads = 3" & Chr(13) & Chr(10)
  writefile f, "SafeTransactions = 0" & Chr(13) & Chr(10)
  writefile f, "PageTimeout = 5" & Chr(13) & Chr(10)
  writefile f, "MaxScanRows = 8" & Chr(13) & Chr(10)
  writefile f, "MaxBufferSize = 2048" & Chr(13) & Chr(10)
  writefile f, "FIL=MS Access" & Chr(13) & Chr(10)
  writefile f, "DriverId = 25" & Chr(13) & Chr(10)
  writefile f, "DefaultDir=" & folder_mdb & Chr(13) & Chr(10)
  writefile f, "DBQ=" & file_mdb & Chr(13) & Chr(10)
  closefile f
End Sub

Sub create_ds_pervasive(nm_db As String)
  Dim f
  
  OpenFile "c:\App_Divisi.dsn", f, 2
  writefile f, "[ODBC]" & Chr(13) & Chr(10)
  writefile f, "DSN=" & nm_db & Chr(13) & Chr(10)
  writefile f, "DRIVER={Pervasive ODBC Engine Interface}" & Chr(13) & Chr(10)
  writefile f, "DBQ=" & nm_db & Chr(13) & Chr(10)
  writefile f, "UID =" & Chr(13) & Chr(10)
  writefile f, "PWD =" & Chr(13) & Chr(10)
  closefile f
End Sub

Public Function week_of(dt As Date, Optional withYear As Boolean = False) As String
     Dim awal_tahun As Date
     Dim wk As String, res1 As String
     
     awal_tahun = "Jan-01-" & Year(dt)
     wk = DateDiff("ww", awal_tahun, dt, vbSunday, vbFirstJan1)
     wk = Trim(str(wk)) + 1  'ini kondisi tahun 2013
     If Len(wk) < 2 Then
          res1 = "0" & Trim(wk)
     Else
          res1 = Trim(str(wk))
     End If
     
     If withYear = True Then
        res1 = Right(Year(dt), 2) & res1
     End If
     
     week_of = res1
End Function




Public Sub CopyFileWindowsWay(SourceFile As String, DestinationFile As String)

     Dim lngReturn As Long
     Dim typFileOperation As SHFILEOPSTRUCT
     

     With typFileOperation
        .hWnd = 0
        .wFunc = FO_COPY
        .pFrom = SourceFile & vbNullChar & vbNullChar 'source file
        .pTo = DestinationFile & vbNullChar & vbNullChar 'destination file
        .fFlags = FOF_ALLOWUNDO
        .fFlags = FOF_NOCONFIRMATION
     End With



     lngReturn = SHFileOperation(typFileOperation)

     If lngReturn <> 0 Then 'Operation failed
          MsgBox Err.LastDllError, vbCritical Or vbOKOnly
     Else 'Aborted
          If typFileOperation.fAnyOperationsAborted = True Then
               MsgBox "Operation Failed", vbCritical Or vbOKOnly
          End If
     End If

End Sub

Function angka2word(angka) As String
  Dim word, word0 As String
  Dim panjang, digit, satuan, puluhan, ratusan As Integer
  Dim c As Long
  Dim tigadigit(5) As String

  If IsNumeric(angka) = False Then
    MsgBox "not valid number"
    angka2word = ""
    Exit Function
  End If
  
  If Trim(angka) = "-" Then
    angka2word = ""
    Exit Function
  End If
  
  word0 = ""
  tigadigit(0) = ""
  tigadigit(1) = " ribu "
  tigadigit(2) = " juta "
  tigadigit(3) = " milyar "
  tigadigit(4) = " trilyun "
  c = 0
  
  Do While Len(RTrim(angka)) > 0
  
  If IsNumeric(angka) = False Then Exit Do
  word = ""
  '3 digit awal
  If Len(RTrim(angka)) > 0 Then
    If IsNumeric(Right(angka, 1)) = True Then
      satuan = CDbl(Right(angka, 1))
    Else
      satuan = 0
    End If
    angka = Left(angka, Len(angka) - 1)
  Else
    satuan = 0
  End If
  If Len(RTrim(angka)) > 0 Then
    If IsNumeric(Right(angka, 1)) = True Then
      puluhan = CDbl(Right(angka, 1))
    Else
      puluhan = 0
    End If
    angka = Left(angka, Len(angka) - 1)
  Else
    puluhan = 0
  End If
  
  'angka satuan dan puluhan
  If puluhan = 1 Then
    Select Case satuan
      Case 0
        word = "sepuluh " & word
      Case 1
        word = "sebelas " & word
      Case 2
        word = "duabelas " & word
      Case 3
        word = "tigabelas " & word
      Case 4
        word = "empatbelas " & word
      Case 5
        word = "limabelas " & word
      Case 6
        word = "enambelas " & word
      Case 7
        word = "tujuhbelas " & word
      Case 8
        word = "delapanbelas " & word
      Case 9
        word = "sembilanbelas " & word
    End Select
  ElseIf puluhan = 0 Then
    Select Case satuan
      Case 0
        If Len(angka) = 0 Then
          word = "nol"
        End If
      Case 1
        word = "satu " & word
      Case 2
        word = "dua " & word
      Case 3
        word = "tiga " & word
      Case 4
        word = "empat " & word
      Case 5
        word = "lima " & word
      Case 6
        word = "enam " & word
      Case 7
        word = "tujuh " & word
      Case 8
        word = "delapan " & word
      Case 9
        word = "sembilan " & word
    End Select
  Else
    Select Case puluhan
      Case 2
        word = "dua puluh " & word
      Case 3
        word = "tiga puluh " & word
      Case 4
        word = "empat puluh " & word
      Case 5
        word = "lima puluh " & word
      Case 6
        word = "enam puluh " & word
      Case 7
        word = "tujuh puluh " & word
      Case 8
        word = "delapan puluh " & word
      Case 9
        word = "sembilan puluh " & word
    End Select
    Select Case satuan
      Case 0
        word = word & ""
      Case 1
        word = word & "satu"
      Case 2
        word = word & "dua"
      Case 3
        word = word & "tiga"
      Case 4
        word = word & "empat"
      Case 5
        word = word & "lima"
      Case 6
        word = word & "enam"
      Case 7
        word = word & "tujuh"
      Case 8
        word = word & "delapan"
      Case 9
        word = word & "sembilan"
    End Select
  End If
  'ratusan
  If Len(RTrim(angka)) > 0 Then
    ratusan = CDbl(Right(angka, 1))
    angka = Left(angka, Len(angka) - 1)
  Else
    ratusan = 0
  End If
  Select Case ratusan
    Case 0
    Case 1
      word = "seratus " & word
    Case 2
      word = "dua ratus " & word
    Case 3
      word = "tiga ratus " & word
    Case 4
      word = "empat ratus " & word
    Case 5
      word = "lima ratus " & word
    Case 6
      word = "enam ratus " & word
    Case 7
      word = "tujuh ratus " & word
    Case 8
      word = "delapan ratus " & word
    Case 9
      word = "sembilan ratus " & word
    End Select
    '----------------------------------
    If Trim(word) <> "" Then
        word = word & tigadigit(c)
    End If
    word0 = word & word0
    c = c + 1
  Loop
  angka2word = word0
End Function

Public Sub pesan2(txt As String, Optional mSecond As Long = 500, _
    Optional warna = vbWhite)
    
  frmPesan.BackColor = warna
  frmPesan.Label1.Caption = txt
  frmPesan.Label1.AutoSize = True
  frmPesan.Width = frmPesan.Label1.Width + 300
  frmPesan.Height = frmPesan.Label1.Height + 300
  
  frmPesan.Label1.Left = Round((frmPesan.Width - frmPesan.Label1.Width) / 2, 0)
  frmPesan.Label1.Top = 150
  
  frmPesan.Timer1.Interval = mSecond
  'frmPesan.Timer1.Enabled = True
  On Error Resume Next
  frmPesan.Show vbModal
End Sub


Function gfIntDB_ADO_OpenDatabase(ByRef prDB_ADOConnection As ADODB.connection, ByVal pvStrConnectionString As String, Optional ByVal pvOptStrUserName As String = "", Optional ByVal pvOptStrPassword As String = "", Optional ByVal pvOptIntDB_ADO_Options As Integer = -1) As Integer
'----------------------------------------------------
' Purpose:
'   Open the ADO Connection to the database
'
' Parmaters:
'   1 - prDB_ADOConnection
'       The database object you wish to open
'   2 - pvStrConnectionString
'       The Pass through ODBC string
'   3 - pvOptStrUserName
'       Optional User Name
'   4 - pvOptStrPassword
'       Optional Password
'   5 - pvOptIntDB_ADO_Options
'       Optional DB Options
'
' Return Value:
'  0  if successful   (gcIntDB_ADO_Opened_Successfully)
' -1  if unsuccessful (gcIntDB_ADO_Opened_Unsuccessfully)
' -2  if unsuccessful (gcIntDB_ADO_Path_or_Pass_Through_Blank)
' -3  if unsuccessful (gcIntDB_ADO_Name_Blank)
' -4  if unsuccessful (gcIntDB_ADO_IndicatorInvalid)
' err if unsuccessful (VB error handler code)
'
'Notes:
' This routine check to see if the database is nothing
' before it attempts to open the database
'
' The database is declared global.
'
'Modification:
'970430  SBI0
'        Added code to correct trailing backslash error
'970425  SBI0
'        Added Parameter to determine query options
'        Added code to close the recordset object if recordset was unsuccessful
'970423  SBI
'        Added check for trailing backslash
'        Added parameter to determine is you want to
'        open the database using access method, or
'        remote server method, such as SQL Server
'
'970225  SBI0
'        New
'-----------------------------------------------------

    Dim lLogContinue        As Integer
    Dim lIntStatusOfOpen    As Integer
    Dim lIntStatusOfClose   As Integer

    Const lcStrBackSlash = "\"
    Const lcIntNoErrors = 0

     'recreate the database connection object
    On Error Resume Next
    Set prDB_ADOConnection = New ADODB.connection
    prDB_ADOConnection = ""
    If Err = lcIntNoErrors Then
        lLogContinue = True
    Else
        lIntStatusOfOpen = Err
    End If
    On Error GoTo 0

    If lLogContinue = True Then

        On Error Resume Next
        'Open the database connection
        prDB_ADOConnection.Open pvStrConnectionString, pvOptStrUserName, pvOptStrPassword, pvOptIntDB_ADO_Options
          DoEvents
        If Err = lcIntNoErrors And prDB_ADOConnection.State = adStateOpen Then
            'return successfull
            lIntStatusOfOpen = 0
        Else
            lIntStatusOfOpen = Err
        End If
        On Error GoTo 0
    End If
    
    gfIntDB_ADO_OpenDatabase = lIntStatusOfOpen
    
End Function
Function gfIntDB_ADO_CloseObject(ByRef prDB_ADO_DatabaseObject As Object) As Integer
'----------------------------------------------------
' Purpose:
'   Closes any DB object
'
' Parmaters:
'   1 - prDB_ADO_DatabaseObject
'       The database object you wish to open
'
' Return Value:
'   0  if successful   (gcIntDB_ADO_Object_Closed_Successfully)
'   -1  if unsuccessful (gcIntDB_ADO_Object_Closed_Unsuccessfully)
'   err if unsuccessful (VB error handler code)
'
' Notes:
'   This routine check to see if the object is nothing
'   before it attempts to close the object
'
' Modification:
'   990204  SBI0
'           Changed to work with ADO
'   970225  SBI0
'           New
'-----------------------------------------------------

    Dim lIntStatusOfObjectClose As Integer
    Const lcIntNoErrors = 0

    On Error Resume Next
    lIntStatusOfObjectClose = -1
    If Not (prDB_ADO_DatabaseObject Is Nothing) Then
        prDB_ADO_DatabaseObject.Close
        Set prDB_ADO_DatabaseObject = Nothing
    End If

    If Err = lcIntNoErrors Then
        lIntStatusOfObjectClose = 0
    Else
        lIntStatusOfObjectClose = Err
    End If
    On Error GoTo 0

    gfIntDB_ADO_CloseObject = lIntStatusOfObjectClose

End Function
Function OpenRecordSet2(ByRef cnn2 As ADODB.connection, ByRef rS2 As ADODB.Recordset, _
        ByVal sql2 As String, Optional ByVal pvOptEnmCursorTypeEnum As CursorTypeEnum = adOpenStatic, _
        Optional ByVal pvOptEnmLockTypeEnum As LockTypeEnum = adLockUnspecified, _
        Optional ByVal curLoc As CursorLocationEnum = adUseServer) As Integer
'----------------------------------------------------
' Purpose:
'   To Run the open recordset function and return a record
'   colletion
'
' Parmaters:
'   4 - pvOptenmCursorTypeEnum
'       Specifies the type of recordset to open
'       adOpenDynamic = 2
'       adOpenForwardOnly = 0
'       adOpenKeyset = 1
'       adOpenStatic = 3  'Routine Default
'       adOpenUnspecified = -1  'VB Default
'   5 - pvOptEnmLockTypeEnum
'       adLockBatchOptimistic = 4
'       adLockOptimistic = 3
'       adLockPessimistic = 2
'       adLockReadOnly = 1
'       adLockUnspecified = -1 'Routine and VB Default
'   6 - pvOptEnmDB_ADO_RecordSetOptions
'       adCmdText           - Provider shouyld take the source as a text description of a command such as a SQL string
'       adCmdTable          - ADO should generate an SQL statement to fetch all rows from the table in Source
'       adCmdTableDirect    - Provider should return all of the rows from the table named in Source
'       adStoredProc        - Provider should treat the Source as a stored procedure
'       AdCmdUnknown        - DO NOT SUE THIS -- SLOWEST of ALL Cursors.  - The command in Source is unknown.
'       adCommandFile       - Saved recordset should be restorted from the file names in the source.
'       asFetchAsync        - Source should be executed asynchronously.
'
' Return Value:
'   0  if successful   (gcIntDB_RecordSet_Opened_Successfully)
'   -1  if unsuccessful (gcIntDB_RecordSet_Opened_UnSuccessfully)
'   -2  if unsuccessful (gcIntDB_RecordSet_Sql_Blank)
'   err if unsuccessful (VB error handler code)
'
' Modification:
'   990216  SBI0
'           Changed from .Execute, to .Open, and added the correct parameters, such as
'           CursorType, Lock Type, and Enum Options.
'   990204  SBI0
'           Changed to work with ADO
'   970420  SBI0
'           Added Parameter to determine query options
'           Added code to close the recordset object if recordset was unsuccessful
'   970225  SBI0
'           New
'-----------------------------------------------------

    Dim lLogContinue                As Integer
    Dim lIntStatusOfRecordSetOpen   As Integer
    Dim lIntStatusOfClose           As Integer
    Const lcIntNoErrors = 0

    lIntStatusOfRecordSetOpen = -1

    'close the open database object
    On Error Resume Next
    Set rS2 = New ADODB.Recordset
    If Err = lcIntNoErrors Then
        lLogContinue = True
    Else
        lIntStatusOfRecordSetOpen = Err
    End If
    
    On Error GoTo 0
    If lLogContinue = True Then
        If sql2 = "" Then
            lLogContinue = False
            lIntStatusOfRecordSetOpen = -2
        End If
    End If

    If lLogContinue = True Then
        On Error Resume Next
        DoEvents
        rS2.CursorLocation = curLoc
        rS2.Open sql2, cnn2, pvOptEnmCursorTypeEnum, pvOptEnmLockTypeEnum
        DoEvents
        If Err = lcIntNoErrors And rS2.State = adStateOpen Then
           lIntStatusOfRecordSetOpen = 0
        Else
            lIntStatusOfRecordSetOpen = Err
'           if the database did not open succesfully
'            make sure recordset object is set to nothing
            lIntStatusOfClose = gfIntDB_ADO_CloseObject(rS2)
        End If
        On Error GoTo 0
    End If
    OpenRecordSet2 = lIntStatusOfRecordSetOpen
End Function

Function OpenRecordSet(ByRef cnn2 As ADODB.connection, _
   ByRef rS2 As ADODB.Recordset, ByVal sql2 As String, _
   Optional ByVal pvOptEnmCursorTypeEnum As CursorTypeEnum = adOpenStatic, _
   Optional ByVal pvOptEnmLockTypeEnum As LockTypeEnum = adLockUnspecified, _
   Optional ByVal curLoc As CursorLocationEnum = adUseServer) As Integer
'----------------------------------------------------
  
  Set rS2 = Nothing
  Set rS2 = New ADODB.Recordset
  
  rS2.CursorType = pvOptEnmCursorTypeEnum
  rS2.LockType = pvOptEnmLockTypeEnum
  rS2.CursorLocation = curLoc
  
  On Error GoTo er1
  DoEvents
  rS2.Open sql2, cnn2
  DoEvents
  OpenRecordSet = 0
  Exit Function
er1:
  MsgBox Err.DESCRIPTION, vbCritical
  Set rS2 = Nothing
  OpenRecordSet = -1
End Function

Function openRs(ByRef cnn As ADODB.connection, ByRef rs As ADODB.Recordset, sql As String) As Integer
    Dim t As Integer
    t = OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient)
    openRs = t
End Function

Function ExecSQL1(ByRef cnn2 As ADODB.connection, _
          sql2 As String) As Integer
    
  'melakukan perintah DML
  'jika 0 maka ok...
  'jika tidak 0, maka error
  
  
  On Error GoTo er1
  cnn2.Execute sql2
  ExecSQL1 = 0
  Exit Function
er1:
  ExecSQL1 = -1
End Function

Function ExecuteSQL(ByRef prDB_ADOConnection As ADODB.connection, ByVal pvStrRecordSource As String) As Integer
'----------------------------------------------------
' Purpose:
'   To execute update, delete and other similar queries
'
'Parmaters:
'   1 - prDB_ADOConnection
'       Database object
'   2 - pvStrRecordSource
'       This is the SQL string
'
' Return Value:
'   0  if successful   (gcIntDB_ADO_Query_Successful)
'   -1  if unsuccessful (gcIntDB_ADO_Query_Unsuccessful)
'   err if unsuccessful (VB error handler code)
'
'
' Notes:
'   This routine will work for both Access and SQLserver
'
'   This routine will run execute queries.
'
'Modification:
'   990204  SBI0
'           Changed to work with ADO
'   970225  SBI0
'           New
'-----------------------------------------------------


    Dim lIntStatusOfQuery As Integer
    Const lcIntNoErrors = 0

    lIntStatusOfQuery = -1

    On Error Resume Next
    DoEvents
    prDB_ADOConnection.Execute pvStrRecordSource
    DoEvents

    If Err = lcIntNoErrors Then
        lIntStatusOfQuery = 0
    Else
        lIntStatusOfQuery = Err
    End If
    On Error GoTo 0

    ExecuteSQL = lIntStatusOfQuery

End Function


Function RecordCount(ByRef prDB_ADORecordset As ADODB.Recordset) As Long
  On Error GoTo er1
  RecordCount = prDB_ADORecordset.RecordCount
  Exit Function
er1:
  RecordCount = -1
End Function

Function RecordCount2(ByRef prDB_ADORecordset As ADODB.Recordset) As Long
  On Error GoTo er1
  prDB_ADORecordset.MoveLast
  RecordCount2 = prDB_ADORecordset.RecordCount
  Exit Function
er1:
  RecordCount2 = -1
End Function

Sub Load_combo(ByRef a As ComboBox, sql As String, _
  ByRef connection As ADODB.connection, _
  Optional set_index As Boolean = True, Optional order = "A", _
  Optional wAll = 0)
  Dim rs As ADODB.Recordset
  
  If OpenRecordSet(connection, rs, sql, adOpenStatic, adLockReadOnly, _
     adUseClient) = 0 Then
    a.Clear
    If wAll = 0 Then
    Else
      a.AddItem "ALL"
    End If
    
    If RecordCount(rs) > 0 Then
      If order = "A" Then
        rs.MoveFirst
        Do While rs.EOF = False
          a.AddItem Trim(cek_null(rs.Fields(0).Value))
          rs.MoveNext
        Loop
      Else
        rs.MoveLast
        Do While rs.BOF = False
          a.AddItem Trim(cek_null(rs.Fields(0).Value))
          rs.MovePrevious
        Loop
      End If
    End If
    If set_index = True Then
      If a.ListCount > 0 Then a.ListIndex = 0
    End If
    Set rs = Nothing
  Else
    MsgBox "Error run " & sql, vbCritical, "Load combo " & a.Name
  End If
End Sub

Sub Load_combo3SQL(ByRef a As ComboBox, sql1 As String, sql2 As String, sql3 As String, _
  ByRef connection As ADODB.connection, _
  Optional set_index As Boolean = True, Optional order = "A", _
  Optional wAll = 0)
  Dim rs As ADODB.Recordset
  
  If OpenRecordSet(connection, rs, sql1, adOpenStatic, adLockReadOnly, _
     adUseClient) = 0 Then
    a.Clear
    If wAll = 0 Then
    Else
      a.AddItem "ALL"
    End If
    
    If RecordCount(rs) > 0 Then
      If order = "A" Then
        rs.MoveFirst
        Do While rs.EOF = False
          a.AddItem Trim(cek_null(rs.Fields(0).Value))
          rs.MoveNext
        Loop
      Else
        rs.MoveLast
        Do While rs.BOF = False
          a.AddItem Trim(rs.Fields(0).Value)
          rs.MovePrevious
        Loop
      End If
    End If
    'If set_index = True Then
    '  If a.ListCount > 0 Then a.ListIndex = 0
    'End If
    Set rs = Nothing
  Else
    MsgBox "Error run Loadcombo3SQL:" & sql1, vbCritical, "Load combo " & a.Name
  End If
  
  '------------- SQL2
  If OpenRecordSet(connection, rs, sql2, adOpenStatic, adLockReadOnly, _
     adUseClient) = 0 Then
        
    If RecordCount(rs) > 0 Then
      If order = "A" Then
        rs.MoveFirst
        Do While rs.EOF = False
          a.AddItem Trim(cek_null(rs.Fields(0).Value))
          rs.MoveNext
        Loop
      Else
        rs.MoveLast
        Do While rs.BOF = False
          a.AddItem Trim(rs.Fields(0).Value)
          rs.MovePrevious
        Loop
      End If
    End If
    Set rs = Nothing
  Else
    MsgBox "Error run Loadcombo3SQL:" & sql2, vbCritical, "Load combo " & a.Name
  End If
  
    '------------- SQL3
  If Trim(sql2) <> "" Then
  If OpenRecordSet(connection, rs, sql3, adOpenStatic, adLockReadOnly, _
     adUseClient) = 0 Then
        
    If RecordCount(rs) > 0 Then
      If order = "A" Then
        rs.MoveFirst
        Do While rs.EOF = False
          a.AddItem Trim(cek_null(rs.Fields(0).Value))
          rs.MoveNext
        Loop
      Else
        rs.MoveLast
        Do While rs.BOF = False
          a.AddItem Trim(rs.Fields(0).Value)
          rs.MovePrevious
        Loop
      End If
    End If
    If set_index = True Then
      If a.ListCount > 0 Then a.ListIndex = 0
    End If
    Set rs = Nothing
  Else
    MsgBox "Error run Loadcombo3SQL:" & sql3, vbCritical, "Load combo " & a.Name
  End If
  End If
End Sub


Sub Load_combo_DiStinct(ByRef a As ComboBox, sql As String, _
  ByRef connection As ADODB.connection, _
  Optional set_index As Boolean = True, Optional order = "A", _
  Optional wAll = 0)
  
  Dim rs As ADODB.Recordset
  Dim rsTemp As New ADODB.Recordset
  Dim item1 As String
  Dim ketemu As Boolean
  Dim c As Long, jRec As Long, c1 As Long
  
  'load list dengan distinct...
  'perintah SQL tanpa distint, semoga bisa lebih cepat
  'isi list ditaruh dulu di rsgrid..
  
  If OpenRecordSet(connection, rs, sql, adOpenStatic, adLockReadOnly, _
     adUseClient) = 0 Then
    a.Clear
    If wAll = 0 Then
    Else
      a.AddItem "ALL"
    End If
    
    jRec = RecordCount(rs)
    If jRec > 0 Then
      
      'create rstemp
      rsTemp.Fields.Append "kd", adVariant
      rsTemp.Open , , adOpenDynamic, adLockPessimistic
      
      If order = "A" Then
        rs.MoveFirst
        Do While rs.EOF = False
          item1 = Trim(cek_null(rs.Fields(0).Value))
          
          'cek, apa item udah ada di rsTemp??
          ketemu = False
          jRec = RecordCount(rsTemp)
          If jRec > 0 Then rsTemp.MoveFirst
          Do While rsTemp.EOF = False
            If Trim(rsTemp(0)) = Trim(item1) Then
              ketemu = True
              Exit Do
            End If
            rsTemp.MoveNext
          Loop
          
          If ketemu = False Then
            rsTemp.AddNew
            rsTemp(0) = item1
            rsTemp.Update
          End If
          rs.MoveNext
        Loop
        
      Else
        rs.MoveLast
        Do While rs.BOF = False
          item1 = Trim(cek_null(rs.Fields(0).Value))
          
          'cek, apa item udah ada di rsTemp??
          ketemu = False
          jRec = RecordCount(rsTemp)
          If jRec > 0 Then rsTemp.MoveFirst
          Do While rsTemp.EOF = False
            If Trim(rsTemp(0)) = Trim(item1) Then
              ketemu = True
              Exit Do
            End If
            rsTemp.MoveNext
          Loop
          If ketemu = False Then
            rsTemp.AddNew
            rsTemp(0) = item1
            rsTemp.Update
          End If
          rs.MovePrevious
        Loop
      End If
    End If
    
      'pindah data dari rsTemp ke combo
        jRec = RecordCount(rsTemp)
        If jRec > 0 Then
          rsTemp.MoveFirst
          Do While rsTemp.EOF = False
            a.AddItem rsTemp(0).Value
            rsTemp.MoveNext
          Loop
        End If
    
    If set_index = True Then
      If a.ListCount > 0 Then a.ListIndex = 0
    End If
    Set rs = Nothing
  Else
    MsgBox "Error run " & sql, vbCritical, "Load combo " & a.Name
  End If
End Sub

Function cari_data1(ByRef connection As ADODB.connection, _
      sql As String, Optional isResAngka As Boolean = False) As String
  
  Dim rs As ADODB.Recordset
  Dim t As String
  
  On Error GoTo er1

  'utk cari data di table, sesuai sql. hasil di kolom pertama
  If OpenRecordSet(connection, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) = 0 Then
    If rs.RecordCount > 0 Then
      If IsNull(rs.Fields(0).Value) = False Then
        t = Trim(rs.Fields(0).Value)
      Else
        t = ""
      End If
    Else
      t = ""
    End If
    
    If isResAngka = True Then
      If Trim(t) = "" Then t = "0"
    End If
    Set rs = Nothing
    cari_data1 = t
  Else
    'MsgBox "Error run " & sql, vbCritical
    sql = InputBox("sql error", "", sql)
    If isResAngka = True Then
      cari_data1 = "-99"
    Else
      cari_data1 = "-99"
    End If
  End If
  Exit Function
er1:
  MsgBox Err.DESCRIPTION, vbCritical
End Function

Sub info(idx, pesan As String, ByRef sb As StatusBar)
  On Error Resume Next
  sb.Panels(idx).text = pesan
  sb.Refresh
  DoEvents
End Sub

Sub info_progress(ByRef sb As StatusBar, idx As Integer, c As Long, jRec As Long, Optional nmProses As String = "")
    Call info(idx, "Load " & nmProses & ". Run " & c & "/" & jRec & " - " & Round((c / jRec) * 100, 2) & "%-", sb)
End Sub


Sub delay(detik As Integer)
  Dim time1 As Date
  Dim time2 As Date
  Dim out As Boolean
  Dim s As Integer
  
  time1 = Now
  out = False
  Do
    time2 = Now
    s = DateDiff("s", time1, time2)
    If s >= detik Then out = True
  Loop While out = False
End Sub

Sub proses(ByRef a As CommandButton, ByRef f As Form, Optional inProgress As Boolean = True)
  If inProgress = True Then
    a.Enabled = False
    f.MousePointer = 11
  Else
    a.Enabled = True
    f.MousePointer = 0
  End If
End Sub

Public Function EncryptString(ByVal InString As String, ByVal EncryptKey As String) As String
 Dim TempKey, OutString As String
 Dim OldChar, NewChar, CryptChar As Long
 Dim i As Integer
 
 ' Initilize i and make sure the EncryptKey is long enough
 i = 0
 Do
  TempKey = TempKey + EncryptKey
 Loop While Len(TempKey) < Len(InString)
 
 ' Loop through the string to encrypt each character.
 Do
  i = i + 1
  OldChar = Asc(Mid(InString, i, 1))
  CryptChar = Asc(Mid(TempKey, i, 1))
  ' If it's an even character, add the ASCII value of the
  ' appropriate character in the Key, otherwise, subract it.
  ' Also, make sure the value is between 0 and 127.
  Select Case i Mod 2
   Case 0      'Even Character
    NewChar = OldChar + CryptChar
    If NewChar > 127 Then NewChar = NewChar - 127
   Case Else   'Odd Character
    NewChar = OldChar - CryptChar
    If NewChar < 0 Then NewChar = NewChar + 127
  End Select
  ' If the value is less than 35, add 40 to it (to make sure
  ' it's in the display range) and put it in an escape
  ' sequence (using ! [ASCII Value 33] as the escape char)
    'MsgBox NewChar
  If NewChar < 35 Then
   OutString = OutString + "!" + Chr(NewChar + 40)
  Else
   OutString = OutString + Chr(NewChar)
  End If
 Loop Until i = Len(InString)
 
 EncryptString = OutString

End Function

Public Function DecryptString(ByVal InString As String, ByVal EncryptKey As String) As String
 Dim TempKey, OutString As String
 Dim OldChar, NewChar, CryptChar As Long
 Dim i, c As Integer
 
 ' Initialize c and i (loop variables)
 c = 0       ' c is used for InString
 i = 0       ' i is used for EncryptKey
 ' Make sure the EncryptKey is long enough
 Do
  TempKey = TempKey + EncryptKey
 Loop While Len(TempKey) < Len(InString)
 
 Do
  ' In the decrypt function, two integers are need keeping
  ' track of location (becuase the escape sequence it two
  ' chars long, but only has one placeholder in the key)
  i = i + 1
  c = c + 1
  OldChar = Asc(Mid(InString, c, 1))
  ' If this is an escape sequence, get the next character and
  ' subtract 40 from it's value.
  If OldChar = 33 Then
   c = c + 1
   OldChar = Asc(Mid(InString, c, 1))
   OldChar = OldChar - 40
  End If
  CryptChar = Asc(Mid(TempKey, i, 1))
  ' If it's an even character, subract the appropriate key
  ' value... also, if it's out of range, bring it back in.
  Select Case i Mod 2
   Case 0      'Even Character
    NewChar = OldChar - CryptChar
    If NewChar < 0 Then NewChar = NewChar + 127
   Case Else   'Odd Character
    NewChar = OldChar + CryptChar
    If NewChar > 127 Then NewChar = NewChar - 127
  End Select
  OutString = OutString + Chr(NewChar)
 Loop Until c = Len(InString)
 
 DecryptString = OutString

End Function

Function cek_sql(ByRef c1 As ADODB.connection, sql As String) As Boolean
  On Error GoTo er1
  c1.Execute sql
  cek_sql = True
  Exit Function
er1:
  cek_sql = False
  'MsgBox Err.Description
End Function

Public Function enkrip(ByVal InString As String, ByVal EncryptKey As String) As String
 Dim TempKey, OutString As String
 Dim OldChar, NewChar, CryptChar As Long
 Dim i As Integer
 
 ' Initilize i and make sure the EncryptKey is long enough
 i = 0
 Do
  TempKey = TempKey + EncryptKey
 Loop While Len(TempKey) < Len(InString)
 
 ' Loop through the string to encrypt each character.
 Do
  i = i + 1
  OldChar = Asc(Mid(InString, i, 1))
  CryptChar = Asc(Mid(TempKey, i, 1))
  ' If it's an even character, add the ASCII value of the
  ' appropriate character in the Key, otherwise, subract it.
  ' Also, make sure the value is between 0 and 127.
  Select Case i Mod 2
   Case 0      'Even Character
    NewChar = OldChar + CryptChar
    If NewChar > 127 Then NewChar = NewChar - 127
   Case Else   'Odd Character
    NewChar = OldChar - CryptChar
    If NewChar < 0 Then NewChar = NewChar + 127
  End Select
  ' If the value is less than 35, add 40 to it (to make sure
  ' it's in the display range) and put it in an escape
  ' sequence (using ! [ASCII Value 33] as the escape char)
  If NewChar < 35 Then
   OutString = OutString + "!" + Chr(NewChar + 40)
  Else
   OutString = OutString + Chr(NewChar)
  End If
 Loop Until i = Len(InString)
 
 enkrip = OutString

End Function

Public Function dekrip(ByVal InString As String, ByVal EncryptKey As String) As String
 Dim TempKey, OutString As String
 Dim OldChar, NewChar, CryptChar As Long
 Dim i, c As Integer
 
 ' Initialize c and i (loop variables)
 c = 0       ' c is used for InString
 i = 0       ' i is used for EncryptKey
 ' Make sure the EncryptKey is long enough
 Do
  TempKey = TempKey + EncryptKey
 Loop While Len(TempKey) < Len(InString)
 
 Do
  ' In the decrypt function, two integers are need keeping
  ' track of location (becuase the escape sequence it two
  ' chars long, but only has one placeholder in the key)
  i = i + 1
  c = c + 1
  
  OldChar = Asc(Mid(InString, c, 1))
  ' If this is an escape sequence, get the next character and
  ' subtract 40 from it's value.
  If OldChar = 33 Then
    c = c + 1
    
    If c > Len(InString) Then Exit Do
    
    OldChar = Asc(Mid(InString, c, 1))
    OldChar = OldChar - 40
  End If
  CryptChar = Asc(Mid(TempKey, i, 1))
  ' If it's an even character, subract the appropriate key
  ' value... also, if it's out of range, bring it back in.
  Select Case i Mod 2
   Case 0      'Even Character
    NewChar = OldChar - CryptChar
    If NewChar < 0 Then NewChar = NewChar + 127
   Case Else   'Odd Character
    NewChar = OldChar + CryptChar
    If NewChar > 127 Then NewChar = NewChar - 127
  End Select
  OutString = OutString + Chr(NewChar)
 Loop Until c >= Len(InString)
 
 dekrip = OutString

End Function
Function get_next_kode(sql As String, ByRef cnn As ADODB.connection, _
   panjang_digit As Integer) As String
   
  'sql mencari kode paling akhir
   
  Dim t As String
  
  t = cari_data1(cnn, sql)
  If Trim(t) = "" Then
    t = "1"
  Else
    t = CStr(CInt(t) + 1)
  End If
  
  Do While Len(Trim(t)) < panjang_digit
    t = "0" & Trim(t)
  Loop
  get_next_kode = t
End Function

Public Function get_autonum(tablename As String, namaKolomId1 As String, _
                            ByRef cn As ADODB.connection) As Long
     'get next number of id1
     Dim sql As String, t As String
     
     sql = "select max(" & Trim(namaKolomId1) & ") from " & Trim(tablename)
     t = cari_data1(cn, sql, True)
     get_autonum = CLng(t) + 1
End Function

Function createRS_duplicate(rsAsli As ADODB.Recordset, ByRef rsCopy As ADODB.Recordset) As Boolean
  'copykan isi dari Rs1 ke RS2
  
  Dim c As Integer
  
  If rsAsli.State <> 1 Then
    createRS_duplicate = False
    Exit Function
  End If
  
  If rsAsli.Fields.Count <= 0 Then
    createRS_duplicate = False
    Exit Function
  End If
  
  Set rsCopy = New ADODB.Recordset
  rsCopy.CursorLocation = adUseClient
  rsCopy.CursorType = adOpenDynamic
  
  For c = 0 To rsAsli.Fields.Count - 1
    rsCopy.Fields.Append rsAsli.Fields(c).Name, adVariant
  Next
  rsCopy.Open , , , adLockPessimistic
  'rsCopy.Open , , adOpenStatic, adLockPessimistic
  createRS_duplicate = True

End Function


Sub create_xls2(rs As ADODB.Recordset, judul As String, kolom_uang As String, _
                kolom_desimal As String, Optional kolom_tanggal As String = "", _
                Optional Kolom_hide As String = "", _
                Optional KetBawah As String = "", Optional cekJmlBaris As Boolean = True, _
                Optional digitKolom As Integer = 2)
  
  'export data dari rs ke excel
  'argumen: kolom_uang, isinya 12,13,10, dst
  'argumen: kolom_desimal, isinya 1,2,3, dst
  'argumen kolom_hide: isinya A,B,H,  dst
  
  Dim baris As Long, kolom As Long
  Dim idx_kolom_akhir As String, klm2, i As Integer, t As String
  Dim t2 As String
  
  On Error GoTo er1
  
  'Dim fl As New Excel.Application
  Dim fl As Object
  Set fl = CreateObject("Excel.Application")
  
  If cekJmlBaris = True Then
  If RecordCount(rs) <= 0 Then
    Call close_xls_lateBinding(fl)
    Exit Sub
  End If
  End If
  
  fl.Workbooks.Add
  
  'read grid
  'header
  
  fl.Range("A1").AddComment
  fl.Range("A1").Comment.Visible = False
  fl.Range("A1").Comment.text ("Judul")
  
  baris = 1
  'judul
  fl.Range("A" & baris & ":AY" & baris).Font.Bold = 2
  fl.Cells(baris, 1) = judul
  If rs.Fields.Count <= 26 Then
    idx_kolom_akhir = Chr(65 + (rs.Fields.Count - 1))
  Else
    idx_kolom_akhir = "Z"
  End If
  fl.Range("A" & baris & ":" & idx_kolom_akhir & baris).Merge
  
  baris = baris + 2
  'nama kolom
  For kolom = 0 To rs.Fields.Count - 1
    fl.Cells(baris, kolom + 1) = UCase(rs.Fields(kolom).Name)
    fl.Range("A" & baris & ":AY" & baris).Font.Bold = 2
  Next
  
  rs.MoveFirst
  baris = baris + 1
  
  Do While rs.EOF = False
    DoEvents
    For kolom = 0 To rs.Fields.Count - 1
      'format utk desimal + duit + tanggal
      If InStr(1, kolom_desimal, CStr(adddigit(kolom, digitKolom)), vbTextCompare) > 0 Then
        'kolom angka
        If IsNumeric(cek_null(rs(kolom), "0")) = True Then
          fl.Cells(baris, kolom + 1) = CDbl(cek_null(rs(kolom), "0"))
        Else
          fl.Cells(baris, kolom + 1) = cek_null(rs(kolom))
        End If
      ElseIf InStr(1, kolom_uang, CStr(adddigit(kolom, digitKolom)), vbTextCompare) > 0 Then
        If IsNumeric(cek_null(rs(kolom), "0")) = True Then
          fl.Cells(baris, kolom + 1) = CCur(cek_null(rs(kolom), "0"))
        Else
          fl.Cells(baris, kolom + 1) = cek_null(rs(kolom))
        End If
      ElseIf InStr(1, kolom_tanggal, CStr(adddigit(kolom, digitKolom)), vbTextCompare) > 0 Then
        t2 = cek_null(rs(kolom))
        If isDate(t2) = True Then
          fl.Cells(baris, kolom + 1) = set_tgl_perv(CDate(t2))
        Else
          fl.Cells(baris, kolom + 1) = cek_null(rs(kolom))
        End If
      ElseIf IsNumeric(cek_null(rs(kolom))) = True Then
          'ditambah petik
          fl.Cells(baris, kolom + 1) = "'" & cek_null(rs(kolom))
      Else
          fl.Cells(baris, kolom + 1) = cek_null(rs(kolom))
      End If
    Next
    rs.MoveNext
    baris = baris + 1
  Loop
  
  'keterangan
  If Trim(KetBawah) <> "" Then
    baris = baris + 1
    fl.Range("A" & baris & ":AY" & baris).Font.Bold = 2
    fl.Cells(baris, 1) = KetBawah
    idx_kolom_akhir = Chr(65 + (rs.Fields.Count - 1))
    fl.Range("A" & baris & ":" & idx_kolom_akhir & baris).Merge
  End If
  
  fl.Columns("A:B").EntireColumn.AutoFit
  fl.Columns("K:M").EntireColumn.AutoFit
  
  If Trim(Kolom_hide) <> "" Then
    klm2 = Split(Kolom_hide, ",", , vbTextCompare)
    For i = 0 To UBound(klm2)
      t = klm2(i)
      On Error Resume Next
      fl.Columns("" & Trim(t) & "").EntireColumn.Hidden = True
    Next
  End If
  
  'fl.DefaultSaveFormat = xlExcel9795
  fl.Workbooks.Close
  
  On Error Resume Next
  Call close_xls_lateBinding(fl)
  
  pesan2 "export ke excel selesai"
  Exit Sub
er1:
  MsgBox Err.DESCRIPTION, vbCritical
  On Error Resume Next
  Call close_xls_lateBinding(fl)
End Sub

Sub create_xls3(rs As ADODB.Recordset, judul As String, kolom_uang As String, _
                kolom_desimal As String, Optional kolom_tanggal As String = "", _
                Optional Kolom_hide As String = "", _
                Optional KetBawah As String = "", Optional cekJmlBaris As Boolean = True, _
                Optional isJudulTampil As Boolean = True, Optional isHeaderTampil As Boolean = True)
  
  'export data dari rs ke excel
  'argumen: kolom_uang, isinya 12,13,10, dst
  'argumen: kolom_desimal, isinya 1,2,3, dst
  'argumen kolom_hide: isinya A,B,H,  dst
  
  Dim baris As Long, kolom As Long
  Dim idx_kolom_akhir As String, klm2, i As Integer, t As String
  Dim t2 As String
  
  On Error GoTo er1
  
  'Dim fl As New Excel.Application
  Dim fl As Object
  Set fl = CreateObject("Excel.Application")
  
  If cekJmlBaris = True Then
  If RecordCount(rs) <= 0 Then
    Call close_xls_lateBinding(fl)
    Exit Sub
  End If
  End If
  
  fl.Workbooks.Add
  
  'read grid
  'header
  
  fl.Range("A1").AddComment
  fl.Range("A1").Comment.Visible = False
  fl.Range("A1").Comment.text ("Judul")
  
  baris = 0
  
  If isJudulTampil = True Then
    'judul
    baris = baris + 1
    fl.Range("A" & baris & ":AY" & baris).Font.Bold = 2
    fl.Cells(baris, 1) = judul
    If rs.Fields.Count <= 26 Then
      idx_kolom_akhir = Chr(65 + (rs.Fields.Count - 1))
    Else
      idx_kolom_akhir = "Z"
    End If
    fl.Range("A" & baris & ":" & idx_kolom_akhir & baris).Merge
    
    baris = baris + 1
  End If
  
  If isHeaderTampil = True Then
    'nama kolom
    baris = baris + 1
    For kolom = 0 To rs.Fields.Count - 1
      fl.Cells(baris, kolom + 1) = UCase(rs.Fields(kolom).Name)
      fl.Range("A" & baris & ":AY" & baris).Font.Bold = 2
    Next
  End If
  
  rs.MoveFirst
  baris = baris + 1
  
  Do While rs.EOF = False
    DoEvents
    For kolom = 0 To rs.Fields.Count - 1
      'format utk desimal + duit + tanggal
      If InStr(1, kolom_desimal, CStr(adddigit(kolom, 2)), vbTextCompare) > 0 Then
        'kolom angka
        If IsNumeric(cek_null(rs(kolom), "0")) = True Then
          fl.Cells(baris, kolom + 1) = CDbl(cek_null(rs(kolom), "0"))
        Else
          fl.Cells(baris, kolom + 1) = cek_null(rs(kolom))
        End If
      ElseIf InStr(1, kolom_uang, CStr(adddigit(kolom, 2)), vbTextCompare) > 0 Then
        If IsNumeric(cek_null(rs(kolom), "0")) = True Then
          fl.Cells(baris, kolom + 1) = CCur(cek_null(rs(kolom), "0"))
        Else
          fl.Cells(baris, kolom + 1) = cek_null(rs(kolom))
        End If
      ElseIf InStr(1, kolom_tanggal, CStr(adddigit(kolom, 2)), vbTextCompare) > 0 Then
        t2 = cek_null(rs(kolom))
        If isDate(t2) = True Then
          fl.Cells(baris, kolom + 1) = set_tgl_perv(CDate(t2))
        Else
          fl.Cells(baris, kolom + 1) = cek_null(rs(kolom))
        End If
      ElseIf IsNumeric(cek_null(rs(kolom))) = True Then
          'ditambah petik
          fl.Cells(baris, kolom + 1) = "'" & cek_null(rs(kolom))
      Else
          fl.Cells(baris, kolom + 1) = cek_null(rs(kolom))
      End If
    Next
    rs.MoveNext
    baris = baris + 1
  Loop
  
  'keterangan
  If Trim(KetBawah) <> "" Then
    baris = baris + 1
    fl.Range("A" & baris & ":AY" & baris).Font.Bold = 2
    fl.Cells(baris, 1) = KetBawah
    idx_kolom_akhir = Chr(65 + (rs.Fields.Count - 1))
    fl.Range("A" & baris & ":" & idx_kolom_akhir & baris).Merge
  End If
  
  fl.Columns("A:B").EntireColumn.AutoFit
  fl.Columns("K:M").EntireColumn.AutoFit
  
  If Trim(Kolom_hide) <> "" Then
    klm2 = Split(Kolom_hide, ",", , vbTextCompare)
    For i = 0 To UBound(klm2)
      t = klm2(i)
      On Error Resume Next
      fl.Columns("" & Trim(t) & "").EntireColumn.Hidden = True
    Next
  End If
  
  'fl.DefaultSaveFormat = xlExcel9795
  fl.Workbooks.Close
  
  On Error Resume Next
  Call close_xls_lateBinding(fl)
  
  pesan2 "export ke excel selesai"
  Exit Sub
er1:
  MsgBox Err.DESCRIPTION, vbCritical
  On Error Resume Next
  Call close_xls_lateBinding(fl)
End Sub

Function open3query(sql1 As String, sql2 As String, sql3 As String, ByRef rs As ADODB.Recordset) As Boolean
         
  'buka satu-satu, tampung di rs
  'sql nya jumlah kolomnya harus sama, nama kolom juga sama
  
  Dim rS2 As ADODB.Recordset
  Dim hasil As Boolean
  Dim a As Integer
  Dim jRec As Long
  
  hasil = True
  
  'buka sql1
  DoEvents
  If OpenRecordSet(cnn, rS2, sql1, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
    DoEvents
    hasil = False
    Exit Function
  End If
  
  'set rs
  Set rs = Nothing
  Set rs = New ADODB.Recordset
  For a = 0 To rS2.Fields.Count - 1
    rs.Fields.Append rS2.Fields(a).Name, adVariant
  Next
  rs.Open , , adOpenDynamic, adLockPessimistic
  
  'pindahkan data rs SQL1 ke rs
  jRec = RecordCount(rS2)
  If jRec > 0 Then
    rS2.MoveFirst
    Do While rS2.EOF = False
       DoEvents
       rs.AddNew
      For a = 0 To rS2.Fields.Count - 1
        rs.Fields(a).Value = rS2.Fields(a).Value
      Next
      rs.Update
      rS2.MoveNext
    Loop
  End If
  
  '--------------------
  
  'buka sql2
  DoEvents
  If OpenRecordSet(cnn, rS2, sql2, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
    DoEvents
    hasil = False
    Exit Function
  End If
  'pindahkan data rs SQL2 ke rs
  jRec = RecordCount(rS2)
  If jRec > 0 Then
    rS2.MoveFirst
    Do While rS2.EOF = False
      DoEvents
      rs.AddNew
      For a = 0 To rS2.Fields.Count - 1
        rs.Fields(a).Value = rS2.Fields(a).Value
      Next
      rs.Update
      rS2.MoveNext
    Loop
  End If
  
  'buka sql3
  DoEvents
  If OpenRecordSet(cnn, rS2, sql3, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
    DoEvents
    hasil = False
    Exit Function
  End If
  'pindahkan data rs SQL3 ke rs
  jRec = RecordCount(rS2)
  If jRec > 0 Then
    rS2.MoveFirst
    Do While rS2.EOF = False
      DoEvents
      rs.AddNew
      For a = 0 To rS2.Fields.Count - 1
        rs.Fields(a).Value = rS2.Fields(a).Value
      Next
      rs.Update
      rS2.MoveNext
    Loop
  End If
  open3query = hasil
  
End Function

Function keKarakterAngka(t As String) As String
  'hilangkan yg bukan angka
  Dim a As Integer
  Dim tBaru As String
  
  tBaru = ""
  For a = 1 To Len(t)
    If IsNumeric(Mid(t, a, 1)) = True Then
      tBaru = tBaru & Mid(t, a, 1)
    End If
  Next
  keKarakterAngka = tBaru
End Function

Function get_string(str1 As String, Charpenutup As String) As String
  'get string sampai batas string yg ditentukan
  
  Dim t2 As String
  Dim a As Integer
  
  t2 = ""
  For a = 1 To Len(str1)
    If Mid(str1, a, 1) <> Trim(Charpenutup) Then
      t2 = t2 & Mid(str1, a, 1)
    Else
      Exit For
    End If
  Next
  get_string = Trim(t2)
End Function

Sub create_table_DDL(ByRef cn As ADODB.connection, sql_Cek As String, _
                    Sql_Create As String, nm_Table As String)
  'cek dulu, apa tabel sudah ada?
  If cek_sql(cn, sql_Cek) = True Then
    'tabel sudah ada
  Else
    pesan2 "Tabel " & nm_Table & " belum ada, akan dicreate"
    If cek_sql(cn, Sql_Create) = True Then
      pesan2 "create tabel " & nm_Table & " sukses"
    Else
        pesan2 "create tabel " & nm_Table & " TIDAK sukses"
    End If
  End If
End Sub

Function is_file_ada(nmFile As String) As Boolean
    Dim fso
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
  If (fso.FileExists(nmFile)) Then
    is_file_ada = True
  Else
    is_file_ada = False
  End If
End Function


Function open_xls_lateBinding(ByRef workBook As Object, nmFile As String) As Integer
    Dim fl2 As Object
    'Dim oWorkbook As Object
    
    On Error GoTo er1
    
    Set fl2 = CreateObject("Excel.Application")
    fl2.Workbooks.Open (nmFile)
    Set workBook = fl2
    'Set oWorkbook = fl2.Workbooks.Open(nmFile)
    'Set open_xls_lateBinding = oWorkbook
    
    open_xls_lateBinding = 0
    Exit Function
er1:
    open_xls_lateBinding = -1
End Function

Sub close_xls_lateBinding(ByRef workBook As Object)
    On Error Resume Next
    'If workBook = True Then
        workBook.Quit
    'End If
    Set workBook = Nothing
End Sub


Function Konvert_angka(t As String) As Double
  'konvert dari angka yang pakai desimal "."
  'ke desimal model indonesia
  'dari stringnya, cari angka ".", masukkan ke bulat, sisanya ke pecahannya
  
  Dim c As Double, h As Integer
  Dim bulat As String, pecahan As String
  
  t = Trim(t)
  If IsNumeric(t) = False Then
    Konvert_angka = 0
    Exit Function
  End If
  
  'jika sudah pakai koma, exit
  If InStr(1, t, ",", vbTextCompare) > 0 Then
    Konvert_angka = CDbl(t)
  End If
  
  h = InStr(1, t, ".", vbTextCompare)
  If h > 0 Then
    bulat = Left(t, h - 1)
    pecahan = Mid(t, h + 1, Len(t))
  Else
    bulat = t
    pecahan = "0"
  End If
  c = CDbl(bulat) + CDbl("0," & pecahan)
  Konvert_angka = c
End Function

Sub tampil_report_pdf(fileReport As String, nmFilePdf As String)
    Dim Appl As New CRAXDRT.Application
    Dim Report As New CRAXDRT.Report
    Set Report = Appl.OpenReport(fileReport)
    
    Report.EnableParameterPrompting = False

    With Report
            .ExportOptions.FormatType = crEFTPortableDocFormat
            .ExportOptions.DestinationType = crEDTDiskFile
            .ExportOptions.DiskFileName = nmFilePdf
           ' location & the file name

            .ExportOptions.PDFExportAllPages = True
            .Export (False)
    End With
End Sub


Sub tampil_report(ByRef CR As CrystalReport, fileReport As String, _
          Optional zoom2 As Integer = 100, Optional select2 As String = "", _
          Optional param1 As String = "")
  
  'param1 : [nama parameter];isinya;true"
  'MsgBox "b1"
  Screen.MousePointer = vbHourglass
  CR.Reset
  CR.WindowShowExportBtn = True
  CR.WindowShowGroupTree = True
  CR.WindowShowPrintBtn = True
  CR.WindowShowCancelBtn = True
  CR.WindowShowPrintSetupBtn = True
  CR.WindowShowRefreshBtn = True
  CR.WindowShowSearchBtn = True
  CR.WindowShowZoomCtl = True
  'MsgBox "b2"
  'MsgBox fileReport
  CR.ReportFileName = fileReport
  'MsgBox "b3"
  If Trim(select2) <> "" Then CR.ReplaceSelectionFormula (select2)
  If Trim(param1) <> "" Then _
    CR.ParameterFields(0) = param1
  CR.RetrieveDataFiles
  CR.WindowState = crptMaximized
  'MsgBox "b4"
  CR.PrintReport
  CR.RetrieveDataFiles
  CR.PageZoom (zoom2)
  Screen.MousePointer = vbDefault
  'MsgBox "b5"
End Sub

Sub db_Access_open(ByRef cn As ADODB.connection, lokasi_db As String, Optional pwd As String = "")
  
  On Error GoTo er1
  Call db_access_close(cn)
  Set cn = New ADODB.connection
  
  cn.ConnectionTimeout = 120
  cn.CommandTimeout = 0
  If Trim(pwd) = "" Then
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Persist Security Info=False;Data Source=" & lokasi_db & _
        ";Mode = readwrite"
    cn.Open
  Else
    'open with password
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.ConnectionString = "Data Source=" & lokasi_db & _
                         ";Mode=Share Deny Read|Share Deny Write;Jet OLEDB:Database Password=" & Trim(pwd)
    cn.CursorLocation = adUseClient
    cn.Open
  End If
  
  Exit Sub
er1:
  MsgBox Err.DESCRIPTION, vbCritical
End Sub


Function db_open(nmFileSpmCabang As String, ByRef cn As ADODB.connection) As Boolean
  'jika true : sukses
  'file database di open, security di buka
  'ketika aplikasi di close, file security di kunci
  
  
  On Error GoTo er1
  
  Call db_close(cn)
  
  'reset security
  Call db_disable_security(nmFileSpmCabang)
  
  
  If db_open_no_pwd(cn, nmFileSpmCabang, 2) = False Then
    db_open = False
  Else
    db_open = True
  End If
    
  Exit Function
er1:
  MsgBox Err.DESCRIPTION, vbCritical
  db_open = False
End Function

Sub db_disable_security(File1 As String)
  Dim filedb As String
  
  filedb = File1
  If is_ada_pwd(filedb) = True Then
    Call RESET_pwd(filedb, "kpu771122")
  End If
End Sub

Sub db_close(ByRef cn As ADODB.connection)
  On Error Resume Next
  cn.Close
  Set cn = Nothing
End Sub

Function is_ada_pwd(file_db As String) As Boolean
  'utk mencek apakah sebuah mdb ada passwordnya ?
  
  Dim str As String
  Dim cn As New ADODB.connection
  
  On Error GoTo er1
    
  Set cn = New ADODB.connection
  'open exlusive to cek password, tanpa diberi inputan password
  str = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & Trim(file_db) & _
        ";Persist Security Info=False"
  
  cn.ConnectionString = str
  cn.CursorLocation = adUseClient
  cn.Open
  cn.Close
  Set cn = Nothing
  is_ada_pwd = False
  Exit Function
er1:
  is_ada_pwd = True
End Function

Sub RESET_pwd(file_db As String, pwd_lama As String)
  Dim str As String
  Dim cn As New ADODB.connection
  
  On Error GoTo er1
  
  'open exlusive to create password
  Set cn = New ADODB.connection
  cn.Provider = "Microsoft.Jet.OLEDB.4.0"
  cn.ConnectionString = "Data Source=" & file_db & _
                         ";Mode=Share Deny Read|Share Deny Write;Jet OLEDB:Database Password=" & Trim(pwd_lama)
  cn.CursorLocation = adUseClient
  cn.Open
  
   
   If cn.mode <> 12 Then
      MsgBox "Your database is not opened exclusively", vbCritical
      Exit Sub
   End If
   
  'alter database password [new pwd] [old pdw]
  str = "ALTER Database Password NULL `" & pwd_lama & "`"

  cn.Execute str
  cn.Close
  Set cn = Nothing
  Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
End Sub

Function db_open_no_pwd(ByRef cn As ADODB.connection, File1 As String, _
        mode1 As Integer) As Boolean
  
  'mode2 : ReadWrite
  'mode1 : read
  
  Dim str1 As String
  
  On Error GoTo er1
  Set cn = Nothing
  Set cn = New ADODB.connection
  
  If mode1 = 2 Then
    str1 = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Persist Security Info=False;Data Source=" & _
          Trim(File1) & ";Mode=ReadWrite"
  Else
    str1 = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Persist Security Info=False;Data Source=" & _
          Trim(File1) & ";Mode=Read"
  End If
  'MsgBox str1
  cn.ConnectionString = str1
  cn.Open
  db_open_no_pwd = True
  Exit Function
er1:
    MsgBox Err.DESCRIPTION
  db_open_no_pwd = False
End Function

Sub db_access_close(ByRef cn As ADODB.connection)
  On Error Resume Next
  cn.Close
  Set cn = Nothing
End Sub


'Sub Load_Excel_2Rs(ByRef fl As Excel.Application, NoSheet As Integer, ByRef rs As ADODB.Recordset _
'                   , ByRef sb As StatusBar, Optional Start_baris As Integer = 1, _
'                   Optional Start_kolom As Integer = 1)
Sub Load_Excel_2Rs(ByRef fl As Object, NoSheet As Integer, ByRef rs As ADODB.Recordset _
                   , ByRef sb As StatusBar, Optional Start_baris As Integer = 1, _
                   Optional start_kolom As Integer = 1)
  'asumsi, data di excel tersusun seperti tabel
  'cari dari 1,1 yang ada isinya,
  'kemudian baca ke kanan, jadikan header, dan create RS nya..
  'kemudian baca data ke bawah
  
  Dim baris As Long, kolom As Long, baris_kosong As Long, a As Long
  'Dim fLs As New Excel.Worksheet
  Dim fLs As Object
  Set fLs = CreateObject("Excel.Application")
  
  Dim t As String, ListKolom As String
  
  Set fLs = fl.Sheets(NoSheet)
  'cari dari baris awal,kolom awal yang ada isinya,
  baris = Start_baris
  kolom = start_kolom
  t = Trim(fLs.Cells(baris, kolom))
  Do While Trim(t) = ""
    baris = baris + 1
    If baris > (Start_baris + 10) Then
      MsgBox "Header Kolom tidak ditemukan", vbCritical
      Set fLs = Nothing
      Call close_xls_lateBinding(fl)
      Exit Sub
    End If
    t = Trim(fLs.Cells(baris, kolom))
  Loop

  'udah ketemu awal baris, baca ke kanan utk baca header+ mbuat rs temp
  Set rs = Nothing
  Set rs = New ADODB.Recordset
  
  t = Trim(fLs.Cells(baris, kolom))
  ListKolom = ""
  Do While Trim(t) <> ""
    '---- add kolom di rs
    t = Replace(t, " ", "_", 1, 1, vbTextCompare)
    t = Replace(t, "-", "_", 1, 1, vbTextCompare)
    
    '-cek dulu, jika nama kolom ini sudah ada??
    Do While InStr(1, ListKolom, t, vbTextCompare) > 0
      t = t & "1"
    Loop
    
    rs.Fields.Append t, adVariant
    ListKolom = ListKolom & " " & t
    
    kolom = kolom + 1
    t = Trim(fLs.Cells(baris, kolom))
  Loop
  rs.CursorLocation = adUseClient
  rs.Open , , adOpenStatic, adLockPessimistic
  
  'baca data, jika kosong, dihitung... sampai max baris_kosong
  baris = baris + 1
  kolom = start_kolom
  baris_kosong = 0
  Do While baris_kosong < 5
    Call info(2, "Load Excel baris " & baris, sb)
    t = Trim(fLs.Cells(baris, kolom))
    If Trim(t) <> "" Then
      baris_kosong = 0
      
      rs.AddNew
      For a = 0 To rs.Fields.Count - 1
        t = Trim(cek_null(fLs.Cells(baris, kolom + a)))
        rs.Fields(a) = cek_null(t)
      Next
      rs.Update
    Else
      baris_kosong = baris_kosong + 1
    End If
    baris = baris + 1
  Loop
  
  
  Set fLs = Nothing
  Call close_xls_lateBinding(fl)
End Sub

Sub format_uang(ByRef a As TextBox)
  'jika bukan numeric, diisi 0
  'jk numeric, diset format uang
  
  If IsNumeric(a.text) = False Then
    a.text = "0"
  Else
    If CCur(a.text) > 0 Then a.text = Format(CCur(a.text), "###,###")
  End If
End Sub

Sub format_Desimal(ByRef a As TextBox)
  'jika bukan numeric, diisi 0
  'jk numeric, diset format desimal
  
  If IsNumeric(a.text) = False Then
    a.text = "0"
  Else
    If CDbl(a.text) > 0 Then a.text = CDbl(a.text)
  End If
End Sub

Function konvert_idx_2KolomExcel(kolom As Integer) As String
  Dim dig1 As Integer, dig2 As Integer
  Dim t As String
  
  'kolom = index kolom
    
  If kolom <= 26 Then
    dig1 = 0
    dig2 = kolom
  Else
    dig1 = kolom \ 26
    dig2 = kolom Mod 26
  End If
  
  If dig1 > 0 And dig2 = 0 Then
   dig1 = dig1 - 1
   dig2 = 26
  End If
  If dig1 > 0 Then
    t = Chr(65 + dig1 - 1) & Chr(65 + dig2 - 1)
  Else
    t = Chr(65 + dig2 - 1)
  End If
  konvert_idx_2KolomExcel = t
End Function

Sub create_rs2(ByRef rs As ADODB.Recordset, kolom As String)
  'susunan kolom dipisahkan dengan semicolon
  'exp: no;nama;alamat;dsb
  
  Dim k
  Dim a As Integer
  
  Set rs = Nothing
  Set rs = New ADODB.Recordset
  
  k = Split(kolom, ";")
  For a = 0 To UBound(k)
    rs.Fields.Append k(a), adVariant
  Next
  'Rs_2.Fields.Append "noreg", adVariant
  'Rs_2.Fields.Append "tgl_JT", adVariant
  rs.CursorLocation = adUseClient
  rs.CursorType = adOpenDynamic
  rs.Open , , adOpenDynamic, adLockPessimistic
End Sub


Sub RS2txt(rs As ADODB.Recordset, Optional namaFile As String = "d:\text1.txt", _
            Optional mode1 = 2, Optional konfirmasi As Boolean = True, _
            Optional pembatas As Boolean = True, Optional pemisah As String = ";", _
            Optional withHeader As Boolean = False, Optional startRec As Long = 1, _
            Optional Jumlah As Long = 0)
  
  'export data dari rs ke txt
  'mode1: 1 = read, 2= write , 8= append
  Dim f
  Dim t As String
  Dim i As Integer, c As Long, cJumlah As Long
  
  On Error GoTo er1
    
  If RecordCount(rs) <= 0 Then
    If pembatas = True Then
        Call OpenFile(namaFile, f, mode1)
        Call writefile(f, "#####" & Chr(13) & Chr(10))
        Call closefile(f)
    End If
    Exit Sub
  End If
  
  rs.MoveFirst
  Call OpenFile(namaFile, f, mode1)
  
  If withHeader = True Then
    t = ""
    For i = 0 To rs.Fields.Count - 1
        If i <> rs.Fields.Count - 1 Then
            t = t & Replace(rs.Fields(i).Name, pemisah, "") & pemisah
        Else
            t = t & Replace(rs.Fields(i).Name, pemisah, "")
        End If
    Next
    t = t & Chr(13) & Chr(10)
    Call writefile(f, t)
  End If
  
  'menulis data
  c = 0
  cJumlah = 0
  Do While rs.EOF = False
    DoEvents
    c = c + 1
    If c < startRec Then
    Else
        cJumlah = cJumlah + 1
        t = ""
        For i = 0 To rs.Fields.Count - 1
            If i <> rs.Fields.Count - 1 Then
                t = t & cleanStr(Replace(rs.Fields(i).Value, pemisah, "")) & pemisah
            Else
                t = t & cleanStr(Replace(rs.Fields(i).Value, pemisah, ""))
            End If
        Next
        t = t & Chr(13) & Chr(10)
        Call writefile(f, t)
        
        If Jumlah > 0 Then
            If cJumlah >= Jumlah Then
                Exit Do
            End If
        End If
    End If
    rs.MoveNext
  Loop
  
  If pembatas = True Then
    Call writefile(f, "#####" & Chr(13) & Chr(10))
  End If
  Call closefile(f)
  If konfirmasi = True Then MsgBox "File telah tersimpan di: " & namaFile, vbInformation
  
  Exit Sub
er1:
  MsgBox Err.DESCRIPTION, vbCritical
  
End Sub

Function cleanNpwp(str As String) As String
    str = cleanStr(str)
    str = Replace(str, "-", "")
    str = Replace(str, ".", "")
    cleanNpwp = str
End Function

Function tbInsert(nmTabel As String, kolom(), isi(), ByRef cn2 As ADODB.connection, _
                Optional idxKey As Integer = -1) As Boolean
                
    Dim sql As String, kolomKey As String, isiKolomKey As String
    Dim c As Integer, t As String
    
    '-- cek jika data sudah ada
    If idxKey <> -1 Then
        For c = 0 To UBound(kolom)
            If c = idxKey Then
                kolomKey = kolom(c)
                isiKolomKey = isi(c)
                Exit For
            End If
        Next
        If kolomKey <> "" Then
            If isDataAda(nmTabel, kolomKey, isiKolomKey, cn2) = True Then
                Call pesan2("data sudah ada", 1)
                tbInsert = False
                Exit Function
            End If
        End If
    End If
    '----------------
    
    'kolom
    t = ""
    For c = 0 To UBound(kolom)
        If c = UBound(kolom) Then
            t = t & kolom(c)
        Else
            t = t & kolom(c) & ", "
        End If
    Next
    '-----
    sql = "insert into " & nmTabel & " (" & t & ") values"
    
    'isi
    t = ""
    For c = 0 To UBound(kolom)
        If c = UBound(kolom) Then
            t = t & "'" & cleanPetik(CStr(isi(c))) & "'"
        Else
            t = t & "'" & cleanPetik(CStr(isi(c))) & "', "
        End If
    Next
    '-----
    
    sql = sql & "(" & t & ")"
    If ExecSQL1(cn2, sql) <> 0 Then
        sql = InputBox("sql error", "", sql)
        tbInsert = False
    Else
        tbInsert = True
    End If
End Function


Function tbUpdate(nmTabel As String, kolom(), isi(), ByRef cn2 As ADODB.connection, _
                 kondisi As String) As Boolean
                
    Dim sql As String, kolomKey As String, isiKolomKey As String
    Dim c As Integer, t As String
        
    'kolom
    
    t = ""
    For c = 0 To UBound(kolom)
        If c = UBound(kolom) Then
            t = t & kolom(c) & " = '" & cleanPetik(CStr(isi(c))) & "' "
        Else
            t = t & kolom(c) & " = '" & cleanPetik(CStr(isi(c))) & "', "
        End If
    Next
    '-----
    sql = "update " & nmTabel & " set " & t & "where " & kondisi
        
    If ExecSQL1(cn2, sql) <> 0 Then
        sql = InputBox("sql error", "", sql)
        tbUpdate = False
    Else
        tbUpdate = True
    End If
End Function

Function tbDelete(nmTabel As String, kolom(), isi(), ByRef cn2 As ADODB.connection) As Boolean
                
    Dim sql As String, kolomKey As String, isiKolomKey As String
    Dim c As Integer, t As String
        
    'kolom
    t = ""
    For c = 0 To UBound(kolom)
        If c = UBound(kolom) Then
            t = t & kolom(c) & " = '" & cleanPetik(CStr(isi(c))) & "' "
        Else
            t = t & kolom(c) & " = '" & cleanPetik(CStr(isi(c))) & "' and "
        End If
    Next
    '-----
    sql = "delete from " & nmTabel & " where " & t
        
    If ExecSQL1(cn2, sql) <> 0 Then
        sql = InputBox("sql error", "", sql)
        tbDelete = False
    Else
        tbDelete = True
    End If
End Function

Function tbGet(ByRef cnn1 As ADODB.connection, namaTabel As String, kolom As String, where1 As String, _
                Optional isResAngka As Boolean = False) As String
    
    'get 1 field from table
    
    Dim sql As String, t As String
    
    sql = "select " & kolom & " from " & namaTabel & " where " & where1
    t = cari_data1(cnn1, sql, isResAngka)
    tbGet = t
End Function

Function cleanStr(str As String, Optional wCleanPetik As Boolean = True) As String
    str = cek_null(str)
    str = Replace(Trim(str), Chr(13) & Chr(10), "")
    str = Replace(str, Chr(10) & Chr(13), "")
    str = Replace(str, vbCr, "")
    str = Replace(str, Chr(13), "")
    str = Replace(str, Chr(10), "")
    str = Replace(str, Chr(34), "")
    str = Replace(str, vbTab, "")
    str = Replace(str, ";", "")
    str = Replace(str, "\", " ")
    
    If wCleanPetik = True Then
        str = cleanPetik(str)
    End If
    
    cleanStr = str
End Function

Function cleanPetik(str As String) As String
    str = Replace(str, "'", "")
    cleanPetik = str
End Function

Sub Load_Csv_2Rs(namaFile As String, ByRef rs As ADODB.Recordset, ByRef sb1 As StatusBar, _
                    Optional karakterPemisah As String = ";", Optional jmlBarisLoad As Integer = 1000)
                    
    Dim f
    Dim t As String, t2 As String
    Dim klm
    Dim a As Integer, jmlKolom As Integer
    Dim c As Long
    
    'On Error GoTo er1
    Call OpenFile(namaFile, f, 1)
    
    t = f.readline
    klm = Split(t, karakterPemisah, , vbTextCompare)
    jmlKolom = UBound(klm)
    
    t2 = ""
    For a = 0 To jmlKolom
        If a = jmlKolom Then
            t2 = t2 & "p" & adddigit(CLng(a), 2)
        Else
            t2 = t2 & "p" & adddigit(CLng(a), 2) & ";"
        End If
        
    Next
    
    Call create_rs2(rs, t2)
            
    c = 1
    Do While Trim(t) <> ""
        If jmlBarisLoad <> 0 Then
            If c > jmlBarisLoad Then Exit Do
        End If
    
        Call info(2, "Read CSV Row : " & c, sb1)
        klm = Split(t, karakterPemisah, , vbTextCompare)
        rs.AddNew
        For a = 0 To jmlKolom
            If a <= UBound(klm) Then
                rs.Fields(a) = cek_null(klm(a))
            Else
                rs.Fields(a) = ""
            End If
        Next
        rs.Update
        On Error GoTo er2
        t = f.readline
        c = c + 1
    Loop
    
    Call pesan2("Load " & namaFile & " selesai")
    
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
    Exit Sub
er2:
    Call pesan2("Load " & namaFile & " selesai")
End Sub

Public Function CekKarakter(TextInput As String, karakterDiHapus As String, Optional setUcase As Boolean = True) As String
  
  Dim h As Integer
  Dim voutput As String

  voutput = Trim(TextInput)
  karakterDiHapus = Trim(karakterDiHapus)
  h = 1
  While h <> 0
    h = InStr(h, voutput, karakterDiHapus, 1)
    If h > 0 Then
      voutput = Left(voutput, h - 1) & Right(voutput, Len(voutput) - h)
      'h = h + 1
    End If
  
  Wend
  If setUcase = True Then voutput = UCase(voutput)
     CekKarakter = voutput
End Function

Sub create_csv2(rs As ADODB.Recordset, namaFile As String, Optional pemisah As String = ";", _
                Optional headerTampil As Boolean = True, Optional headerTambahan As String = "", _
                Optional jikaKosong As String = "", Optional indexKolomIsiBlank As String = "")
  
  'export data dari rs ke csv
  'index kolom dari 0
  
  Dim f
  Dim baris1 As String, noBaris As Integer
  Dim kolom As Integer
  Dim temp1 As String, idxKolom As String
  
  On Error GoTo er1
  
  
  If RecordCount(rs) <= 0 Then
    Exit Sub
  End If
  
  'open file text, write
  OpenFile namaFile, f, 2
  
  'header tambahan
  If Trim(headerTambahan) = "" Then
  Else
    writefile f, headerTambahan & Chr(13) & Chr(10)
  End If

  'tulis header dulu
  If headerTampil = True Then
    baris1 = ""
    For kolom = 0 To rs.Fields.Count - 1
        
        If kolom = rs.Fields.Count - 1 Then
            baris1 = baris1 & CekKarakter(rs(kolom).Name, pemisah, False)
        Else
            baris1 = baris1 & CekKarakter(rs(kolom).Name, pemisah, False) & pemisah
        End If
    Next
    writefile f, baris1 & Chr(13) & Chr(10)
  End If
  
  rs.MoveFirst
  noBaris = 1
  Do While rs.EOF = False
    DoEvents
    baris1 = ""
    For kolom = 0 To rs.Fields.Count - 1
        
        'cek untuk kolom yang default blank
        idxKolom = adddigit(CLng(kolom), 2)
        If InStr(1, indexKolomIsiBlank, idxKolom, vbTextCompare) > 0 Then
            If kolom = rs.Fields.Count - 1 Then
                baris1 = baris1 & cleanStr(CekKarakter(cek_null(rs(kolom)), pemisah, False))
            Else
                baris1 = baris1 & cleanStr(CekKarakter(cek_null(rs(kolom)), pemisah, False)) & pemisah
            End If
        Else
            If Trim(jikaKosong) <> "" Then
                'jika kosong ada default isinya...
                temp1 = cleanStr(CekKarakter(cek_null(rs(kolom)), pemisah, False))
                If Trim(temp1) = "" Then
                    temp1 = jikaKosong
                End If
            
                If kolom = rs.Fields.Count - 1 Then
                    baris1 = baris1 & temp1
                Else
                    baris1 = baris1 & temp1 & pemisah
                End If
            Else
                'If kolom = rs.Fields.Count - 1 Then
                '    baris1 = baris1 & cleanStr(CekKarakter(cek_null(rs(kolom)), pemisah, False))
                'Else
                    baris1 = baris1 & cleanStr(CekKarakter(cek_null(rs(kolom)), pemisah, False)) & pemisah
                'End If
            End If
        End If
    Next
    noBaris = noBaris + 1
    If Right(baris1, 1) = pemisah Then
        baris1 = Left(baris1, Len(baris1) - 1)
    End If
    writefile f, baris1 & Chr(13) & Chr(10)
    rs.MoveNext
  Loop
  
  
  On Error Resume Next
  Call closefile(f)
  
  pesan2 "export ke CSV selesai." & vbCr & "File di simpan di: " & namaFile, 1000
  Exit Sub
er1:
  MsgBox Err.DESCRIPTION, vbCritical
  On Error Resume Next
  Call closefile(f)
End Sub

Sub create_csv(rs As ADODB.Recordset, namaFile As String, Optional pemisah As String = ";", _
                Optional headerTampil As Boolean = True, Optional headerTambahan As String = "", _
                Optional jikaKosong As String = "", Optional indexKolomIsiBlank As String = "")
  
  'export data dari rs ke csv
  'index kolom dari 0
  
  Dim f
  Dim baris1 As String
  Dim kolom As Integer
  Dim temp1 As String, idxKolom As String
  
  On Error GoTo er1
  
  
  If RecordCount(rs) <= 0 Then
    Exit Sub
  End If
  
  'open file text, write
  OpenFile namaFile, f, 2
  
  'header tambahan
  If Trim(headerTambahan) = "" Then
  Else
    writefile f, headerTambahan & Chr(13) & Chr(10)
  End If

  'tulis header dulu
  If headerTampil = True Then
    baris1 = ""
    For kolom = 0 To rs.Fields.Count - 1
        
        If kolom = rs.Fields.Count - 1 Then
            baris1 = baris1 & CekKarakter(rs(kolom).Name, pemisah, False)
        Else
            baris1 = baris1 & CekKarakter(rs(kolom).Name, pemisah, False) & pemisah
        End If
    Next
    writefile f, baris1 & Chr(13) & Chr(10)
  End If
  
  rs.MoveFirst
  Do While rs.EOF = False
    DoEvents
    baris1 = ""
    For kolom = 0 To rs.Fields.Count - 1
        
        'cek untuk kolom yang default blank
        idxKolom = adddigit(CLng(kolom), 2)
        If InStr(1, indexKolomIsiBlank, idxKolom, vbTextCompare) > 0 Then
            If kolom = rs.Fields.Count - 1 Then
                baris1 = baris1 & cleanStr(CekKarakter(cek_null(rs(kolom)), pemisah, False))
            Else
                baris1 = baris1 & cleanStr(CekKarakter(cek_null(rs(kolom)), pemisah, False)) & pemisah
            End If
        Else
            If Trim(jikaKosong) <> "" Then
                'jika kosong ada default isinya...
                temp1 = cleanStr(CekKarakter(cek_null(rs(kolom)), pemisah, False))
                If Trim(temp1) = "" Then
                    temp1 = jikaKosong
                End If
            
                If kolom = rs.Fields.Count - 1 Then
                    baris1 = baris1 & temp1
                Else
                    baris1 = baris1 & temp1 & pemisah
                End If
            Else
                If kolom = rs.Fields.Count - 1 Then
                    baris1 = baris1 & cleanStr(CekKarakter(cek_null(rs(kolom)), pemisah, False))
                Else
                    baris1 = baris1 & cleanStr(CekKarakter(cek_null(rs(kolom)), pemisah, False)) & pemisah
                End If
            End If
        End If
        
    Next
    If Right(baris1, 1) = pemisah Then
        baris1 = Left(baris1, Len(baris1) - 1)
    End If
    writefile f, baris1 & Chr(13) & Chr(10)
    rs.MoveNext
  Loop
  
  
  On Error Resume Next
  Call closefile(f)
  
  pesan2 "export ke CSV selesai." & vbCr & "File di simpan di: " & namaFile, 1000
  Exit Sub
er1:
  MsgBox Err.DESCRIPTION, vbCritical
  On Error Resume Next
  Call closefile(f)
End Sub

Sub selAllText(ByRef tb1 As TextBox)
    tb1.SelLength = Len(tb1.text)
End Sub

Function NearestThousand(num As Currency) As Currency
    Dim t As String
    t = CStr(num)
    t = Mid(t, 1, Len(t) - 3)
    
    NearestThousand = CCur(t) * 1000
End Function



Sub SendMessage(MailFrom, MailTo, Subject, Message)
    Dim ObjSendMail
    Set ObjSendMail = CreateObject("CDO.Message")

    'This section provides the configuration information for the remote SMTP server.

    With ObjSendMail.Configuration.Fields
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2  'Send the message using the network (SMTP over the network).
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smpt server Address"
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False  'Use SSL for the connection (True or False)
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

    ' If your server requires outgoing authentication uncomment the lines below and use a valid email address and password.
'    .Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
'    .Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = MailFrom
'    .Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = yourpassword

    .Update
    End With

    'End remote SMTP server configuration section==

    ObjSendMail.To = MailTo
    ObjSendMail.Subject = Subject
    ObjSendMail.From = MailFrom
    ObjSendMail.AddAttachment "c:\mydocuments\test.txt"

    ' we are sending a html email.. simply switch the comments around to send a text email instead
    ObjSendMail.HTMLBody = Message
    'ObjSendMail.TextBody = Message

    ObjSendMail.Send

    Set ObjSendMail = Nothing
End Sub

Function createRS_duplicateContent(rsAsli As ADODB.Recordset, ByRef rsCopy As ADODB.Recordset, _
                                    ByRef sb1 As StatusBar, _
                                    Optional createRsCopy As Boolean = True) As Boolean
  'copykan kolom dari Rs1 ke RS2
  
  Dim c As Long, jRec As Long
  Dim d As Integer
  
    If createRsCopy = True Then
        If createRS_duplicate(rsAsli, rsCopy) = False Then
            createRS_duplicateContent = False
            Exit Function
        End If
    End If
    
    jRec = RecordCount(rsAsli)
    If jRec <= 0 Then
        createRS_duplicateContent = True
        Exit Function
    End If
    
    
    rsAsli.MoveFirst
    c = 1
    Do While rsAsli.EOF = False
        Call info_progress(sb1, 1, c, jRec, "Copy RS")
        
        rsCopy.AddNew
        
        For d = 0 To rsAsli.Fields.Count - 1
            rsCopy.Fields(d).Value = rsAsli.Fields(d).Value
        Next
        
        
        rsAsli.MoveNext
        c = c + 1
    Loop
    rsCopy.Update
    createRS_duplicateContent = True
End Function
