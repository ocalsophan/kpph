Attribute VB_Name = "modFtp"
Option Explicit

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" _
         (ByVal dwAccess As Long, _
         ByVal fInherit As Integer, _
         ByVal hObject As Long) As Long
         
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const SYNCHRONIZE = 1048576


Public Function DownloadFile(ftpServer As String, uname As String, pwd As String) As Boolean
 Dim strMsg$
 Dim strLine$
 Dim gstrMessage$
 Dim processID#
 Dim process_handle#
 
 Dim strDwnldFilePth As String
 Dim strdwnldbtchpth As String
 
 '------------------------ Authentication
 
 '
 ' Use your own path but App.Path & "\strDwnldFilePth.txt will work
 '
 'strDwnldFilePth = App.Path & "\sftp.txt"
 strDwnldFilePth = "c:\sftp.txt"
 Open strDwnldFilePth For Output As #1
 
 '--
 'uname
 'password
 'mode ascii, dst
 '"cd /outgoing/articles/" & vbCrLf & _

 strLine = ""
 strLine = uname & vbCrLf & _
           pwd & vbCrLf & _
           "ascii" & vbCrLf & _
           "prompt off" & vbCrLf & _
           "mget *.html" & vbCrLf & _
           "close" & vbCrLf & _
           "bye" & vbCrLf

 Print #1, strLine
 Close #1

 '
 ' Use your own path but App.Path & "\strdwnldbtchpth.bat will work
 '
 strdwnldbtchpth = App.Path & "\strdwnldbtchpth.bat"
 Open strdwnldbtchpth For Output As #1
 
 strLine = ""

 '
 ' Use your own path instead of \projectFTPData or create this folder on your C drive
 ' Also, App.Path & "\strDwnldFilePth.txt will work
 '
 strLine = "C:" & vbCrLf & _
           "cd " & "\projectFTPData\" & vbCrLf & _
           "C:\WINDOWS\system32\ftp.exe -s:" & "c:\sftp.txt" & " " & ftpServer & vbCrLf

 Print #1, strLine
 Close #1

 '----------------------- Downloading Data with FTP ------------------------------
 
 '
 ' Use your own path but App.Path & "\strdwnldbtchpth.bat will work
 '
 processID = Shell(App.Path & "\strdwnldbtchpth.bat", vbHide)
 process_handle = OpenProcess(SYNCHRONIZE, 0, processID)

 'NOT NEEDED--->If process_handle <> 0 Then
 WaitForSingleObject process_handle, INFINITE
 CloseHandle process_handle
 'NOT NEEDED--->End If

 gstrMessage = gstrMessage & "Searching for files in <IN> folder" & vbCrLf
 'gstrMessageToUser = gstrMessageToUser & "Searching for files in <IN> folder" & vbCrLf

 '
 ' Use your own path instead of \projectFTPData or create this folder on your C drive
 '
 If Dir$("c:\projectFTPData" & "\ART*.ZIP") = "" Then
   strMsg = strMsg & "No files on FTP to download" & vbCrLf
   MsgBox strMsg
   'Fn_DownloadFiles = False
   'flgFileDownloaded = False
   Exit Function
 Else
   DownloadFile = True
   'flgFileDownloaded = True
   strMsg = strMsg & "Files Downloaded from FTP" & vbCrLf
   MsgBox strMsg
 End If
End Function


Public Function UploadFile(ftpServer As String, uname As String, pwd As String, _
                        port As String, namaFile As String) As Boolean
    Dim strLine$
    
    Dim strDwnldFilePth As String
    Dim strdwnldbtchpth As String
 
    strDwnldFilePth = App.Path & "\exp\sftp.txt"
    Open strDwnldFilePth For Output As #1
    
    'ini script perintah winscp, utk upload
    strLine = ""
    strLine = "echo on" & vbCrLf & _
                "open -hostkey=""*"" sftp://" & uname & ":" & pwd & "@" & ftpServer & " -rawsettings" & vbCrLf & _
                "cd /Upload/Users/TAXPAYMENT" & vbCrLf & _
                "lcd """ & App.Path & "\exp\""" & vbCrLf & _
                "mput " & """" & namaFile & """" & vbCrLf & _
                "close" & vbCrLf & _
                "bye" & vbCrLf

    Print #1, strLine
    Close #1

 
    'ini script bat yang akan di panggil
    strdwnldbtchpth = App.Path & "\strdwnldbtchpth.bat"
    Open strdwnldbtchpth For Output As #1
 
    strLine = ""
    strLine = """c:\Program Files (x86)\WinSCP\WinSCP.com"" /script=""" & _
                App.Path & "\exp\sftp.txt""" & vbCrLf & _
            "pause" & vbCrLf
    strLine = """" & tbVariabel_get("winscp") & """ /script=""" & _
                App.Path & "\exp\sftp.txt""" & vbCrLf & _
            "pause" & vbCrLf

    Print #1, strLine
    Close #1

    'panggil file batnya
    Call Shell(App.Path & "\strdwnldbtchpth.bat", vbNormalFocus)
 
End Function

Public Function DownloadFile2(ftpServer As String, uname As String, pwd As String, _
                        port As String, namaFile As String) As Boolean
    Dim strLine$
    
    Dim strDwnldFilePth As String
    Dim strdwnldbtchpth As String
 
    strDwnldFilePth = App.Path & "\exp\sftp.txt"
    Open strDwnldFilePth For Output As #1
    
    'ini script perintah winscp, utk upload
    strLine = ""
    strLine = "echo on" & vbCrLf & _
                "open -hostkey=""*"" sftp://" & uname & ":" & pwd & "@" & ftpServer & vbCrLf & _
                "cd /Downloads" & vbCrLf & _
                "lcd """ & App.Path & "\exp\""" & vbCrLf & _
                "get *.*" & vbCrLf & _
                "close" & vbCrLf & _
                "bye" & vbCrLf

    Print #1, strLine
    Close #1

 
    'ini script bat yang akan di panggil
    strdwnldbtchpth = App.Path & "\strdwnldbtchpth.bat"
    Open strdwnldbtchpth For Output As #1
 
    strLine = ""
    strLine = """" & tbVariabel_get("winscp") & """ /script=""" & _
                App.Path & "\exp\sftp.txt""" & vbCrLf & _
            "pause" & vbCrLf

    Print #1, strLine
    Close #1

    'panggil file batnya
    Call Shell(App.Path & "\strdwnldbtchpth.bat", vbNormalFocus)
 
End Function

