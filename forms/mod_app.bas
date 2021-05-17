Attribute VB_Name = "mod_app"
Option Explicit


Function checkNPWP(input1 As String) As Boolean
    'numeric only, pastikan npwp hanya angka
    Dim npwp As String
    Dim ret1 As Boolean
    Dim padnumber(0 To 9) As String
    Dim Total1 As Integer, ceilUp As Integer, validationNumber As Integer, validationResult As Integer
    Dim serialNumber As String
    Dim validationSerialNumber As Boolean
    Dim c As Integer
    Dim t As String, kpp As String, sql As String
    Dim kppCode()
    
    npwp = cekStringAngka(input1)

    'make sure it's 15 number
    If Len(npwp) <> 15 Then
        checkNPWP = False
        Exit Function
    End If
    'multiply factor
    Dim multiplyby
    multiplyby = Array(1, 2, 1, 2, 1, 2, 1, 2)

    'find first 8 digits
    
    serialNumber = Left(npwp, 8)

    'first 8 digit multiply by "multiply factor, hasilnya 2 digit
    For c = 1 To 8
        t = CStr(CInt(Mid(serialNumber, c, 1)) * multiplyby(c - 1))
        padnumber(c) = adddigit(CLng(t), 2)
    Next
    
    'total digit kiri dan kanan
    Total1 = 0
    For c = 1 To 8
        t = padnumber(c)
        Total1 = Total1 + CInt(Mid(t, 1, 1)) + CInt(Mid(t, 2, 1))
    Next
    
    'ceil up total to nearest 10
    'ceilUp = Math.ceil(Total / 10) * 10
    ceilUp = Round((Total1 / 10) + 0.5) * 10

    'validation code or number, 9th character
    validationNumber = CInt(Mid(npwp, 9, 1))

    'validation number after calculation
    validationResult = ceilUp - Total1
    
    If validationResult = 10 Then validationResult = 0
    
    'validation of serial number
    If validationNumber = validationResult Then
        validationSerialNumber = True
    Else
        validationSerialNumber = False
        checkNPWP = validationSerialNumber
        Exit Function
    End If

    'KPP code
   kppCode = Array("000", "001", "002", "003", "004", "005", "006", "007", "008", "009", "010", "011", "012", "013", "014", "015", "016", "017", "018", "019", _
        "020", "021", "022", "023", "024", "025", "026", "027", "028", "029", "030", "031", "032", "033", "034", "035", "036", "037", "038", "039", _
        "040", "041", "042", "043", "044", "045", "046", "047", "048", "050", "051", "052", "053", "054", "055", "056", "057", "058", "059", "060", _
        "061", "062", "063", "064", "065", "066", "067", "070", "071", "072", "073", "074", "075", "076", "077", "080", "081", "085", "086", "090", _
        "091", "092", "093", "100", "101", "102", "103", "104", "105", "106", "107", "110", "111", "112", "113", "114", "115", "116", "117", "118", _
        "119", "120", "121", "122", "123", "124", "125", "126", "127", "128", "130", "140", "150", "160", "170", "180", "190", "200", "201", "202", _
        "203", "204", "205", "210", "211", "212", "213", "214", "215", "216", "217", "218", "219", "220", "221", "222", "223", "224", "225", "230", _
        "240", "250", "260", "270", "280", "290", "300", "301", "302", "303", "304", "305", "306", "307", "308", "309", "310", "311", "312", "313", _
        "314", "315", "320", "321", "322", "323", "324", "325", "326", "327", "328", "330", "331", "332", "333", "334", "335", "401", "402", "403", _
        "404", "405", "406", "407", "408", "409", "411", "412", "413", "414", "415", "416", "417", "418", "419", "421", "422", "423", "424", "425", _
        "426", "427", "428", "429", "431", "432", "433", "434", "435", "436", "437", "438", "439", "441", "442", "443", "444", "445", "446", "447", _
        "448", "451", "452", "453", "501", "502", "503", "504", "505", "506", "507", "508", "509", "511", "512", "513", "514", "515", "516", "517", _
        "518", "521", "522", "523", "524", "525", "526", "527", "528", "529", "531", "532", "533", "541", "542", "543", "544", "545", "601", "602", _
        "603", "604", "605", "606", "607", "608", "609", "611", "612", "613", "614", "615", "616", "617", "618", "619", "621", "622", "623", "624", _
        "625", "626", "627", "628", "629", "631", "641", "642", "643", "644", "645", "646", "647", "648", "649", "651", "652", "653", "654", "655", _
        "656", "657", "701", "702", "703", "704", "705", "706", "711", "712", "713", "714", "721", "722", "723", "724", "725", "726", "727", "728", _
        "729", "731", "732", "733", "734", "735", "801", "802", "803", "804", "805", "806", "807", "808", "809", "811", "812", "813", "814", "815", _
        "816", "821", "822", "823", "824", "825", "831", "832", "833", "834", "901", "902", "903", "904", "905", "906", "907", "908", "911", "912", _
        "913", "914", "915", "921", "922", "923", "924", "925", "926", "941", "942", "943", "951", "952", "953", "954", "955", "956", "454")
    
    kpp = Mid(npwp, 10, 3)
    'If InStrArray(kpp, kppCode) = True Then
    '    validationSerialNumber = True
    'Else
    '    validationSerialNumber = False
    'End If
    
    sql = "select count(*) from mkppCode where code1 = '" & kpp & "'"
    t = cari_data1(cnn, sql, True)
    If CInt(t) > 0 Then
        validationSerialNumber = True
    Else
        validationSerialNumber = False
    End If
    checkNPWP = validationSerialNumber

End Function



Sub update_tabel_temp()
    Dim sql As String
       
    'open cnnTemp
    Call db_Access_open(cnnTemp, App.Path & "\data\dbrep.mdb")
    
    '-- tabel bp_pph23
    sql = "CREATE TABLE bp_pph23 ( " & _
        "NPWP_KPP char(30) , " & _
        "kd_proyek char(30) , " & _
        "nott char(30) , " & _
        "nofaktur char(30) , " & _
        "Kode_Form char(20) , " & _
        "Masa_Pajak char(5) , " & _
        "Tahun_Pajak char(10) , " & _
        "Pembetulan char(10) , " & _
        "NPWP_WP char(50) , " & _
        "Nama_WP char(100) , " & _
        "Alamat_WP char(100) , " & _
        "Nomor_Bukti_Potong char(100) , " & _
        "Tanggal_Bukti_Potong DATE , " & _
        "Nilai_Bruto_1 currency , " & _
        "Tarif_1 char(10) , " & _
        "PPh_Yang_Dipotong__1 currency  , " & _
        "Nilai_Bruto_2 currency , " & _
        "Tarif_2 char(10) , " & _
        "PPh_Yang_Dipotong__2 currency , " & _
        "Nilai_Bruto_3 currency , " & _
        "Tarif_3 char(10)  , " & _
        "PPh_Yang_Dipotong__3 currency , " & _
        "Nilai_Bruto_4 currency , " & _
        "Tarif_4 char(10) , "
    sql = sql & "PPh_Yang_Dipotong__4 currency , " & _
        "Nilai_Bruto_5 currency  , " & _
        "Tarif_5 char(10) , " & _
        "PPh_Yang_Dipotong__5 currency , " & _
        "Nilai_Bruto_6a currency , " & _
        "Tarif_6a char(10) , " & _
        "PPh_Yang_Dipotong__6a currency , " & _
        "Nilai_Bruto_6b currency , " & _
        "Tarif_6b char(10) , " & _
        "PPh_Yang_Dipotong__6b currency , " & _
        "Nilai_Bruto_6c currency , " & _
        "Tarif_6c char(10) , " & _
        "PPh_Yang_Dipotong__6c currency , " & _
        "Nilai_Bruto_9 currency , " & _
        "Tarif_9 char(10) , " & _
        "PPh_Yang_Dipotong__9 currency , " & _
        "Nilai_Bruto_10 currency , " & _
        "Perkiraan_Penghasilan_Netto10 currency , " & _
        "Tarif_10 char(10) , " & _
        "PPh_Yang_Dipotong__10 currency , " & _
        "Nilai_Bruto_11 currency , " & _
        "Perkiraan_Penghasilan_Netto11 currency , " & _
        "Tarif_11 char(10) , " & _
        "PPh_Yang_Dipotong__11 currency , " & _
        "Nilai_Bruto_12 currency , "
   sql = sql & "Perkiraan_Penghasilan_Netto12 currency , " & _
        "Tarif_12 char(10)  , " & _
        "PPh_Yang_Dipotong__12 currency , " & _
        "Nilai_Bruto_13 currency , " & _
        "Tarif_13 currency , " & _
        "PPh_Yang_Dipotong__13 currency , " & _
        "Kode_Jasa_6d1 char(30) , " & _
        "Nilai_Bruto_6d1 currency , " & _
        "Tarif_6d1 char(10) , " & _
        "PPh_Yang_Dipotong__6d1 currency , " & _
        "Kode_Jasa_6d2 char(30) , " & _
        "Nilai_Bruto_6d2 currency , " & _
        "Tarif_6d2 char(10) , " & _
        "PPh_Yang_Dipotong__6d2 currency , " & _
        "Kode_Jasa_6d3 char(30) , " & _
        "Nilai_Bruto_6d3 currency , " & _
        "Tarif_6d3 char(10) , " & _
        "PPh_Yang_Dipotong__6d3 currency , " & _
        "Kode_Jasa_6d4 char(30) , " & _
        "Nilai_Bruto_6d4 currency , " & _
        "Tarif_6d4 char(10) , " & _
        "PPh_Yang_Dipotong__6d4 currency , " & _
        "Kode_Jasa_6d5 char(30) , " & _
        "Nilai_Bruto_6d5 currency , " & _
        "Tarif_6d5 char(10) , "
   sql = sql & "PPh_Yang_Dipotong__6d5 currency , " & _
        "Kode_Jasa_6d6 char(30) , " & _
        "Nilai_Bruto_6d6 currency , " & _
        "Tarif_6d6 char(10) , " & _
        "PPh_Yang_Dipotong__6d6 currency , " & _
        "Jumlah_Nilai_Bruto_ currency , " & _
        "Jumlah_PPh_Yang_Dipotong currency , " & _
        "kode_divisi char(10) , " & _
        "tgl_import DATE , " & _
        "id1 long NOT NULL, " & _
        "email char(255), " & _
        "Primary Key(id1)" & _
        ")"
    Call create_table_DDL(cnnTemp, "select * from bp_pph23", sql, "bp_pph23")
    
    '-- tabel bp_pph22
    sql = "CREATE TABLE bp_pph22 (" & _
        "NPWP_KPP char(30) , " & _
        "kd_proyek char(30) , " & _
        "nott char(30) , " & _
        "nofaktur char(30) , " & _
        "k02 char(30)  , " & _
        "Masa_Pajak char(5) , " & _
        "Tahun_Pajak char(10) , " & _
        "Pembetulan char(10) , " & _
        "NPWP char(50) , " & _
        "Nama_NPWP char(100) , " & _
        "Alamat char(100) , " & _
        "Nomor_Bukti_Potong char(100) , " & _
        "Tanggal_Bukti_Potong DATE , " & _
        "k11 char(30) , " & _
        "k12 char(30) , " & _
        "k13 char(30) , " & _
        "k14 char(30) , " & _
        "k15 char(30) , " & _
        "k16 char(30) , "
    sql = sql & "k17 char(30) , " & _
        "k18 char(30) , " & _
        "k19 char(30) , " & _
        "k20 char(30) , " & _
        "k21 char(30) , " & _
        "k22 char(30) , " & _
        "k23 char(30) , " & _
        "k24 char(30) , " & _
        "k25 char(30) , " & _
        "k26 char(30) , " & _
        "k27 char(30) , " & _
        "k28 char(30) , " & _
        "k29 char(30) , " & _
        "k30 char(30) , " & _
        "k31 char(30) , " & _
        "k32 char(30) , " & _
        "k33 char(30) , " & _
        "k34 char(30) , " & _
        "k35 char(30) , " & _
        "k36 char(30) , "
    sql = sql & "k37 char(30) , " & _
        "k38 char(30) , " & _
        "k39 char(30) , " & _
        "k40 char(30) , " & _
        "k41 char(30) , " & _
        "k42 char(30) , " & _
        "k43 char(30) , " & _
        "Nilai_DPP currency, " & _
        "Tarif char(10) , " & _
        "Nilai_PPh currency, " & _
        "k47 char(30) , " & _
        "k48 char(30) , " & _
        "k49 char(30) , " & _
        "k50 char(30) , " & _
        "j51 char(30) , " & _
        "j52 char(30) , " & _
        "kode_divisi char(10), " & _
        "tgl_import DATE , " & _
        "id1 long , " & _
        "email char(255), " & _
        "Primary Key(id1) " & _
        ")"
    Call create_table_DDL(cnnTemp, "select * from bp_pph22", sql, "bp_pph22")
    
    
    '-- tabel bp_pph42_konstruksi
    sql = "CREATE TABLE bp_pph42_konstruksi ( " & _
        "NPWP_KPP char(30), " & _
        "kd_proyek char(30) , " & _
        "nott char(30) , " & _
        "nofaktur char(30) , " & _
        "Kode_Form char(30) , " & _
        "Masa_Pajak char(10) , " & _
        "Tahun_Pajak char(10) , " & _
        "Pembetulan char(10) , " & _
        "NPWP_WP char(50), " & _
        "Nama_WP char(100) , " & _
        "Alamat_WP char(100) , " & _
        "Nomor_Bukti_Potong char(100) , " & _
        "Tanggal_Bukti_Potong DATE , " & _
        "Jenis_Hadiah_Undian_1 char(50), " & _
        "Kode_Option_Tempat_Penyimpanan_1 char(30) , " & _
        "Jumlah_Nilai_Bruto_1 currency, " & _
        "Tarif_1 char(30) , " & _
        "PPh_Yang_Dipotong__1 currency, " & _
        "Jenis_Hadiah_Undian_2 char(30), "
    sql = sql & "Kode_Option_Tempat_Penyimpanan_2 char(30) , " & _
        "Jumlah_Nilai_Bruto_2 currency, " & _
        "Tarif_2 char(30) , " & _
        "PPh_Yang_Dipotong__2 currency, " & _
        "Jenis_Hadiah_Undian_3 char(30), " & _
        "Kode_Option_Tempat_Penyimpanan_3 char(30) , " & _
        "Jumlah_Nilai_Bruto_3 currency, " & _
        "Tarif_3 char(30) , " & _
        "PPh_Yang_Dipotong__3 currency, " & _
        "Jenis_Hadiah_Undian_4 char(30), " & _
        "Kode_Option_Tempat_Penyimpanan_4 char(30) , " & _
        "Jumlah_Nilai_Bruto_4 currency, " & _
        "Tarif_4 char(30) , " & _
        "PPh_Yang_Dipotong__4 currency, " & _
        "Jenis_Hadiah_Undian_5 char(30), " & _
        "Kode_Option_Tempat_Penyimpanan_5 char(30) , " & _
        "Jumlah_Nilai_Bruto_5 currency, " & _
        "Tarif_5 char(30) , " & _
        "PPh_Yang_Dipotong__5 currency, " & _
        "Jenis_Hadiah_Undian_6 char(30), "
    sql = sql & "Jumlah_Nilai_Bruto_6 currency, " & _
        "Tarif_6 char(30) , " & _
        "PPh_Yang_Dipotong__6 currency, " & _
        "Jumlah_Nilai_Bruto_7 currency, " & _
        "Tarif_7 char(30), " & _
        "PPh_Yang_Dipotong_7 currency, " & _
        "Jenis_Penghasilan_8 char(30) , " & _
        "Jumlah_Nilai_Bruto_8 currency, " & _
        "Tarif_8 char(30) , " & _
        "PPh_Yang_Dipotong_8 currency, " & _
        "Jumlah_PPh_Yang_Dipotong currency, " & _
        "Tanggal_Jatuh_Tempo_Obligasi char(50) , " & _
        "Tanggal_Perolehan_Obligasi char(50) , " & _
        "Tanggal_Penjualan_Obligasi char(50) , " & _
        "Holding_Periode_Obligasi char(30), " & _
        "Time_Periode_Obligasi char(30), " & _
        "kode_divisi char(10) , " & _
        "tgl_import DATE, " & _
        "id1 long, " & _
        "email char(255), "
    sql = sql & "Primary Key(id1) " & _
        ")"
    'sql = InputBox("", "", sql)
    Call create_table_DDL(cnnTemp, "select * from bp_pph42_konstruksi", sql, "bp_pph42_konstruksi")
    
    
    '-- tabel bp_pph42_sewa
    sql = "CREATE TABLE bp_pph42_sewa ( " & _
        "NPWP_KPP char(30), " & _
        "kd_proyek char(30) , " & _
        "nott char(30) , " & _
        "nofaktur char(30) , " & _
        "Kode_Form char(30) , " & _
        "Masa_Pajak char(10) , " & _
        "Tahun_Pajak char(10) , " & _
        "Pembetulan char(10) , " & _
        "NPWP_WP char(50), " & _
        "Nama_WP char(100) , " & _
        "Alamat_WP char(100) , " & _
        "Nomor_Bukti_Potong char(100) , " & _
        "Tanggal_Bukti_Potong DATE , " & _
        "Jenis_Hadiah_Undian_1 char(50), " & _
        "Kode_Option_Tempat_Penyimpanan_1 char(30) , " & _
        "Jumlah_Nilai_Bruto_1 currency, " & _
        "Tarif_1 char(30) , " & _
        "PPh_Yang_Dipotong__1 currency, " & _
        "Jenis_Hadiah_Undian_2 char(30), "
    sql = sql & "Kode_Option_Tempat_Penyimpanan_2 char(30) , " & _
        "Jumlah_Nilai_Bruto_2 currency, " & _
        "Tarif_2 char(30) , " & _
        "PPh_Yang_Dipotong__2 currency, " & _
        "Jenis_Hadiah_Undian_3 char(30), " & _
        "Kode_Option_Tempat_Penyimpanan_3 char(30) , " & _
        "Jumlah_Nilai_Bruto_3 currency, " & _
        "Tarif_3 char(30) , " & _
        "PPh_Yang_Dipotong__3 currency, " & _
        "Jenis_Hadiah_Undian_4 char(30), " & _
        "Kode_Option_Tempat_Penyimpanan_4 char(30) , " & _
        "Jumlah_Nilai_Bruto_4 currency, " & _
        "Tarif_4 char(30) , " & _
        "PPh_Yang_Dipotong__4 currency, " & _
        "Jenis_Hadiah_Undian_5 char(30), " & _
        "Kode_Option_Tempat_Penyimpanan_5 char(30) , " & _
        "Jumlah_Nilai_Bruto_5 currency, " & _
        "Tarif_5 char(30) , " & _
        "PPh_Yang_Dipotong__5 currency, " & _
        "Jenis_Hadiah_Undian_6 char(30), "
    sql = sql & "Jumlah_Nilai_Bruto_6 currency, " & _
        "Tarif_6 char(30) , " & _
        "PPh_Yang_Dipotong__6 currency, " & _
        "Jumlah_Nilai_Bruto_7 currency, " & _
        "Tarif_7 char(30), " & _
        "PPh_Yang_Dipotong_7 currency, " & _
        "Jenis_Penghasilan_8 char(30) , " & _
        "Jumlah_Nilai_Bruto_8 currency, " & _
        "Tarif_8 char(30) , " & _
        "PPh_Yang_Dipotong_8 currency, " & _
        "Jumlah_PPh_Yang_Dipotong currency, " & _
        "Tanggal_Jatuh_Tempo_Obligasi char(50) , " & _
        "Tanggal_Perolehan_Obligasi char(50) , " & _
        "Tanggal_Penjualan_Obligasi char(50) , " & _
        "Holding_Periode_Obligasi char(30), " & _
        "Time_Periode_Obligasi char(30), " & _
        "kode_divisi char(10) , " & _
        "tgl_import DATE, " & _
        "id1 long, " & _
        "email char(255), "
    sql = sql & "Primary Key(id1) " & _
        ")"
    Call create_table_DDL(cnnTemp, "select * from bp_pph42_sewa", sql, "bp_pph42_sewa")
    
    '------------
    'perubahan di kolom pph22
    sql = "alter table pph22  " & _
        "add namakpp char(100), " & _
        "kotakpp char(100);"
    Call create_table_DDL(cnnTemp, "select namakpp from pph22", sql, "update pph22 namakpp")
    
    'perubahan di kolom pph23
    sql = "alter table pph23  " & _
        "add namakpp char(100), " & _
        "kotakpp char(100);"
    Call create_table_DDL(cnnTemp, "select namakpp from pph23", sql, "update pph23 namakpp")
    
    'perubahan di kolom pph42_konstruksi
    sql = "alter table pph42_konstruksi  " & _
        "add namakpp char(100), " & _
        "kotakpp char(100);"
    Call create_table_DDL(cnnTemp, "select namakpp from pph42_konstruksi", sql, "update pph42_konstruksi namakpp")
    
    'perubahan di kolom pph42_sewa
    sql = "alter table pph42_sewa  " & _
        "add namakpp char(100), " & _
        "kotakpp char(100);"
    Call create_table_DDL(cnnTemp, "select namakpp from pph42_sewa", sql, "update pph42_sewa namakpp")
    '----------
    
    
    '-- update table
    'perubahan di kolom pph22
    sql = "alter table pph22  " & _
        "add email char(255), " & _
        "terbilang char(255);"
    Call create_table_DDL(cnnTemp, "select terbilang from pph22", sql, "update pph22")
    
    'perubahan di kolom pph23
    sql = "alter table pph23  " & _
        "add email char(255), " & _
        "terbilang char(255);"
    Call create_table_DDL(cnnTemp, "select terbilang from pph23", sql, "update pph23")
    
    'perubahan di kolom pph42_konstruksi
    sql = "alter table pph42_konstruksi  " & _
        "add email char(255), " & _
        "terbilang char(255);"
    Call create_table_DDL(cnnTemp, "select terbilang from pph42_konstruksi", sql, "update pph42_konstruksi")
    
    'perubahan di kolom pph42_sewa
    sql = "alter table pph42_sewa  " & _
        "add email char(255), " & _
        "terbilang char(255);"
    Call create_table_DDL(cnnTemp, "select terbilang from pph42_sewa", sql, "update pph42_sewa")
    'close koneksi
    Call db_access_close(cnnTemp)
End Sub



Function format_Npwp_awal(t As String) As String
    Dim hasil As String
    
    '09.321.683.6
    t = Trim(t)
    hasil = Left(t, 2) & "." & Mid(t, 3, 3) & "." & Mid(t, 6, 3) & "." & Mid(t, 9, 1)
    format_Npwp_awal = hasil
End Function

Sub cek_npwpWP(npwp_wp As String, nama As String, alamat As String)
    'cek, jika npwp_WP belum ada, insert
    If isDataAda("mnpwp", "npwp", npwp_wp, cnn, False) = True Then
        'data sudah ada
    Else
        Call tbMNpwp_insert(npwp_wp, nama, alamat, 0)
    End If
    
End Sub

Sub create_ds_mySQL()
  Dim f
  Dim file_DSN As String
  
  file_DSN = "c:\dbpph.dsn"
  OpenFile file_DSN, f, 2
  writefile f, "[ODBC]" & Chr(13) & Chr(10)
  writefile f, "DRIVER=MySQL ODBC 5.1 Driver" & Chr(13) & Chr(10)
  writefile f, "UID=trunojoy_rep" & Chr(13) & Chr(10)
  writefile f, "PWD=urep2017" & Chr(13) & Chr(10)
  writefile f, "PORT=3306" & Chr(13) & Chr(10)
  writefile f, "DATABASE=trunojoy_dbpph" & Chr(13) & Chr(10)
  writefile f, "SERVER=trunojoyopython.com" & Chr(13) & Chr(10)
  closefile f
End Sub

Function lokasi_server_load(namaFile As String) As String
    'cek, file adakah ? jika tidak ada, create default value, terus di load
    'file \data\set_db.txt
    
    Dim f
    Dim namaServer As String
    
    On Error GoTo er1
    If is_file_ada(namaFile) = False Then
        'create file
        Call OpenFile(namaFile, f, 2)
        Call writefile(f, "trunojoyopython.com")
        Call closefile(f)
    End If
    
    'load file
    Call OpenFile(namaFile, f, 1)
    namaServer = f.readline
    Call closefile(f)
    lokasi_server_load = namaServer
    Exit Function
er1:
    MsgBox Err.DESCRIPTION, vbCritical
    lokasi_server_load = ""
End Function

Sub lokasi_server_save(namaFile As String, namaServer As String)
    Dim f
    
    'create file
    Call OpenFile(namaFile, f, 2)
    Call writefile(f, namaServer)
    Call closefile(f)
    Call pesan2("Lokasi Server di Update", 1, vbYellow)
End Sub

Function cek_Koneksi() As Boolean
    On Error GoTo err_DoWebRequest
    Dim strurl As String, DoWebRequest As String
    strurl = "http://www.google.com"
    
    Dim objXML As Object
    Set objXML = CreateObject("Microsoft.XMLHTTP")
    objXML.Open "GET", strurl, False
    objXML.Send
    If (objXML.Status = 404) Then
        DoWebRequest = "404 Error"
        DoWebRequest = objXML.responseText
        cek_Koneksi = False
    Else
        cek_Koneksi = True
    End If
    Set objXML = Nothing
    Exit Function
err_DoWebRequest:
    cek_Koneksi = False
End Function

Sub load_Divisi(ByRef cb As ComboBox, Optional load_filter As Boolean = False, _
                 Optional wAll As Integer = 0, Optional setIndex As Boolean = False)
     
     Dim pesan As String, sql  As String
     Dim Label1 As String
     
     Label1 = "Divisi "
     pesan = ""
     If load_filter = True Then
          pesan = InputBox("Filter " & Label1, "Filter " & Label1, "")
     End If
     
     If Trim(pesan) = "" Then
          sql = "select distinct concat(kodedivisi,' - ',nama_divisi) from mdivisi order by kodedivisi limit 20"
     Else
          sql = "select distinct concat(kodedivisi,' - ',nama_divisi) from mdivisi " & _
                "where ucase(kodeDivisi) like '%" & UCase(Trim(pesan)) & _
                    "%' or ucase(namaDivisi) like '%" & UCase(Trim(pesan)) & "%' order by kodedivisi limit 20 "
     End If
     Call Load_combo(cb, sql, cnn, setIndex, , wAll)
End Sub

Sub load_Proyek(kd_Divisi As String, jenisPajak As String, ByRef cb As ComboBox, Optional load_filter As Boolean = False, _
                 Optional wAll As Integer = 1, Optional setIndex As Boolean = False)
     
     Dim pesan As String, sql  As String
     Dim Label1 As String
     
     Label1 = "Proyek "
     pesan = ""
     If load_filter = True Then
          pesan = InputBox("Filter " & Label1, "Filter " & Label1, "")
     End If
     
     If Trim(pesan) = "" Then
          sql = "select distinct kd_proyek from " & jenisPajak & " where kode_divisi = '" & Trim(kd_Divisi) & "' order by kd_proyek"
     Else
          sql = "select distinct kd_proyek from " & jenisPajak & _
                "where kode_divisi = '" & Trim(kd_Divisi) & "' and ucase(kd_proyek) like '%" & UCase(Trim(pesan)) & "%' order by kd_proyek"
     End If
     Call Load_combo(cb, sql, cnn, setIndex, , wAll)
End Sub

Sub load_KPP(ByRef cb As ComboBox, Optional load_filter As Boolean = False, _
                 Optional wAll As Integer = 0, Optional setIndex As Boolean = True)
     
     Dim pesan As String, sql  As String
     Dim Label1 As String
     
     Label1 = "KPP "
     pesan = ""
     If load_filter = True Then
          pesan = InputBox("Filter " & Label1, "Filter " & Label1, "")
     End If
     
     If Trim(pesan) = "" Then
          sql = "select distinct concat(npwp,' # ', kpp_administrasi) from mkpp order by kpp_administrasi limit 20"
     Else
          sql = "select distinct concat(npwp,' # ', kpp_administrasi) from mkpp " & _
                "where ucase(npwp) like '%" & UCase(Trim(pesan)) & _
                    "%' or ucase(kpp_administrasi) like '%" & UCase(Trim(pesan)) & "%' order by kpp_administrasi limit 20 "
     End If
     
     
     Call Load_combo(cb, sql, cnn, setIndex, , wAll)
End Sub

Sub load_Tahun_pph15(ByRef cb As ComboBox)
     
    Dim sql  As String
     
    sql = "select distinct tahun_pajak from pph15"
    Call Load_combo(cb, sql, cnn, True, , 1)
End Sub

Sub load_Masa_pph15(ByRef cb As ComboBox)
     
    Dim sql  As String
     
    sql = "select distinct masa_pajak from pph15"
    Call Load_combo(cb, sql, cnn, True, , 1)
End Sub

Sub load_Tahun2(ByRef cb As ComboBox, nmTabel1 As String)
     
    Dim sql  As String
     
    sql = "select distinct Tahun_Pajak from " & nmTabel1
    Call Load_combo(cb, sql, cnn, True, , 1)
End Sub

Sub load_TahunEkualisasi(ByRef cb As ComboBox)
     
    Dim sql  As String
     
    sql = "select distinct Tahun from tbaccpac"
    Call Load_combo(cb, sql, cnn, True, , 0)
End Sub

Sub load_Masa2(ByRef cb As ComboBox, nmTabel1 As String)
     
    Dim sql  As String
     
    sql = "select distinct Masa_Pajak from " & nmTabel1
    Call Load_combo(cb, sql, cnn, True, , 1)
End Sub

Sub load_Pembetulan2(ByRef cb As ComboBox, nmTabel1 As String)
     
    Dim sql  As String
     
    sql = "select distinct Pembetulan from " & nmTabel1
    Call Load_combo(cb, sql, cnn, True, , 1)
End Sub



Sub load_jenisPPh(ByRef cb As ComboBox)
     
     Dim Label1 As String
     
     cb.Clear
     cb.AddItem "1. PPh ps15"
     cb.AddItem "2. PPh ps23"
     cb.AddItem "3. PPh ps21 Tidak Final"
     cb.AddItem "4. PPh ps21 Bulanan"
     cb.AddItem "5. PPh ps21 Tahunan"
     cb.AddItem "6. PPh ps22"
     cb.AddItem "7. PPh ps26"
     cb.AddItem "8. PPh ps4 ayat2 Konstruksi"
     cb.AddItem "9. PPh ps4 ayat2 Sewa"
     cb.AddItem "10. PPh 4(2) Bunga Obligasi"
     cb.AddItem "11. PPh 21 Dibawah PTKP"
     cb.AddItem "12. PPh 21 Pesangon Final"
     
     'cb.ListIndex = 0
End Sub

Sub load_jenisPPhSsp(ByRef cb As ComboBox)
     
     Dim Label1 As String
     
     cb.Clear
     cb.AddItem "ALL"
     cb.AddItem "PPh Pasal 21"
     cb.AddItem "PPh Pasal 22"
     cb.AddItem "PPh Pasal 23"
     cb.AddItem "PPh Final"
     cb.AddItem "PPh Pasal 15"
     
     'cb.ListIndex = 0
End Sub

Function tbM_Ptkp_getNilai(ptkp As String) As Currency
    Dim sql As String, t As String
    
    sql = "select nilai from mptkp where key1 = '" & Trim(ptkp) & "'"
    t = cari_data1(cnn, sql, True)
    tbM_Ptkp_getNilai = CCur(t)
End Function

Function get_pph21Setahun(pkp_Setahun As Currency) As Currency
    '=IF(P5<50000001;ROUND((5/100*P5);0);IF(P5<250000001;ROUND(2500000+(15/100*(P5-50000000));0);
    'IF(P5<500000001;ROUND(32500000+(25/100*(P5-250000000));0);
    'IF(P5>500000000;ROUND(95000000+(30/100*(P5-500000000));0);0))))
    
    Dim t As Currency
    
    If pkp_Setahun < 50000001 Then
        t = Round((5 / 100 * pkp_Setahun), 0)
    ElseIf pkp_Setahun < 250000001 Then
        t = Round(2500000 + (15 / 100 * (pkp_Setahun - 50000000)), 0)
    ElseIf pkp_Setahun < 500000001 Then
        t = Round(32500000 + (25 / 100 * (pkp_Setahun - 250000000)), 0)
    ElseIf pkp_Setahun > 500000000 Then
        t = Round(95000000 + (30 / 100 * (pkp_Setahun - 500000000)), 0)
    End If
    get_pph21Setahun = t
End Function

Function tbM_Ptkp_isDataAda(key1 As String) As Boolean
    If isDataAda("mptkp", "key1", key1, cnn) = True Then
        tbM_Ptkp_isDataAda = True
    Else
        tbM_Ptkp_isDataAda = False
    End If
End Function

Function tbM_Ptkp_insert(key1 As String, nilai As Currency) As Boolean
    Dim sql As String
    
    If isDataAda("mptkp", "key1", key1, cnn) = True Then
        Call pesan2("Kode " & key1 & " sudah ada", , vbYellow)
        tbM_Ptkp_insert = False
        Exit Function
    Else
        sql = "insert into mptkp (key1, nilai) values ('" & key1 & "','" & nilai & "')"
        If ExecSQL1(cnn, sql) <> 0 Then
            MsgBox "error run " & sql, vbInformation
            tbM_Ptkp_insert = False
        Else
            tbM_Ptkp_insert = True
        End If
    End If
End Function

Function tbM_Ptkp_Delete(key1 As String) As Boolean
    Dim sql As String
    
    sql = "delete from mptkp where key1 = '" & cleanStr(key1) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        MsgBox "error run " & sql, vbInformation
        tbM_Ptkp_Delete = False
    Else
        tbM_Ptkp_Delete = True
    End If
End Function

Function tbmkppCode_Delete(code1 As String) As Boolean
    Dim sql As String
    
    sql = "delete from mkppCode where code1 = '" & cleanStr(code1) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        MsgBox "error run " & sql, vbInformation
        tbmkppCode_Delete = False
    Else
        tbmkppCode_Delete = True
    End If
End Function

Function tbmkppCode_insert(code1 As String) As Boolean
    Dim sql As String
    
    If isDataAda("mkppCode", "code1", code1, cnn) = True Then
        Call pesan2("Kode " & code1 & " sudah ada", , vbYellow)
        tbmkppCode_insert = False
        Exit Function
    Else
        sql = "insert into mkppCode (code1) values ('" & _
                cleanStr(code1) & "')"
        If ExecSQL1(cnn, sql) <> 0 Then
            MsgBox "error run " & sql, vbInformation
            tbmkppCode_insert = False
        Else
            tbmkppCode_insert = True
        End If
    End If
End Function

Function tbMKpp_insert(npwp As String, nama As String, alamat As String, tgl_lahir As Date, klu As String, _
                        nip_nama_ar As String, status_update As String, tgl_update As Date, _
                        kpp_administrasi As String) As Boolean
    Dim sql As String
    
    If isDataAda("mkpp", "npwp", npwp, cnn) = True Then
        Call pesan2("Kode " & npwp & " sudah ada", , vbYellow)
        tbMKpp_insert = False
        Exit Function
    Else
        If Trim(kpp_administrasi) = "" Then
            Call pesan2("Data tidak lengkap", , vbYellow)
            tbMKpp_insert = False
            Exit Function
        End If
        
        npwp = Replace(npwp, ".", "")
        npwp = Replace(npwp, "-", "")
        
        sql = "insert into mkpp (npwp, nama, alamat, " & _
                "tgl_lahir, klu, nip_nama_ar, " & _
                "status_update, tgl_update, kpp_administrasi) values ('" & _
                cleanStr(npwp) & "','" & cleanStr(nama) & "','" & cleanStr(alamat) & "','" & _
                set_tgl_perv(tgl_lahir) & "','" & cleanStr(klu) & "','" & cleanStr(nip_nama_ar) & "','" & _
                cleanStr(status_update) & "','" & set_tgl_perv(tgl_update) & "','" & cleanStr(kpp_administrasi) & "')"
        If ExecSQL1(cnn, sql) <> 0 Then
            MsgBox "error run " & sql, vbInformation
            tbMKpp_insert = False
        Else
            tbMKpp_insert = True
        End If
    End If
End Function

Function tbMKpp_Update(npwp As String, nama As String, alamat As String, tgl_lahir As Date, klu As String, _
                        nip_nama_ar As String, status_update As String, tgl_update As Date, _
                        kpp_administrasi As String) As Boolean
    Dim sql As String
    
    sql = "update mkpp set nama = '" & cleanStr(nama, True) & "', alamat = '" & _
            cleanStr(alamat, True) & "', tgl_lahir = '" & set_tgl_perv(tgl_lahir) & "', klu = '" & _
            cleanStr(klu, True) & "', nip_nama_ar = '" & cleanStr(nip_nama_ar, True) & "', status_update = '" & _
            cleanStr(status_update, True) & "', tgl_update = '" & _
            set_tgl_perv(tgl_update) & "', kpp_administrasi = '" & cleanStr(kpp_administrasi) & "' where npwp = '" & _
            cleanStr(npwp, True) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        MsgBox "error run " & sql, vbInformation
        tbMKpp_Update = False
    Else
        tbMKpp_Update = True
    End If
End Function

Function tbMKpp_isNpwpKPP_Valid(npwp_kpp As String) As Boolean
    Dim sql As String, t As String
    
    npwp_kpp = Replace(npwp_kpp, ".", "")
    npwp_kpp = Replace(npwp_kpp, "-", "")
    
    sql = "select count(*) from mkpp where npwp = '" & cleanStr(npwp_kpp) & "'"
    t = cari_data1(cnn, sql, True)
    If CInt(t) > 0 Then
        tbMKpp_isNpwpKPP_Valid = True
    Else
        tbMKpp_isNpwpKPP_Valid = False
    End If
End Function

Function tbMKpp_Delete(npwp As String) As Boolean
    Dim sql As String
    
    sql = "delete from mkpp where npwp = '" & cleanStr(npwp) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        MsgBox "error run " & sql, vbInformation
        tbMKpp_Delete = False
    Else
        tbMKpp_Delete = True
    End If
End Function

Function tbMKpp_get(kolom As String, npwp_kpp) As String
    Dim sql As String, t As String
    
    sql = "select " & kolom & " from mkpp where npwp = '" & Trim(npwp_kpp) & "'"
    t = cari_data1(cnn, sql)
    tbMKpp_get = t
End Function

Function tbMKpp_getNamaKPP(npwp_kpp As String) As String
    Dim res1 As String
    
    res1 = tbMKpp_get("kpp_administrasi", npwp_kpp)
    If IsNumeric(Left(res1, 3)) = True Then
        res1 = Trim(Mid(res1, 7, 100))
    End If
    tbMKpp_getNamaKPP = res1
End Function

Function tbMKpp_getKotaKPP(npwp_kpp As String) As String
    Dim res1 As String
    Dim kota As String
    
    res1 = UCase(Trim(tbMKpp_getNamaKPP(npwp_kpp)))
    If res1 = "WAJIB PAJAK BESAR EMPAT" Then
        kota = "JAKARTA"
    Else
        res1 = Replace(res1, "PRATAMA", "")
        res1 = Replace(res1, "MADYA", "")
        kota = Trim(res1)
    End If
    tbMKpp_getKotaKPP = kota
End Function


Function dbMySQL_close()
    'close connecttion
    On Error Resume Next
    cnn.Close
    Set cnn = Nothing
End Function

Function dbMySQL_open() As Boolean
    'open
    Dim cnnString As String, namaServer As String
        
    
    DoEvents
    
    'If IsOnline = False Then
    '    Call pesan2("Internet tidak aktif", 1)
    '    dbMySQL_open = False
    '    Exit Function
    'End If
    
    'If Not cnn Is Nothing Then
    '    If cnn.State = adStateOpen Then
    '        'masih ok
    '        dbMySQL_open = True
    '        Exit Function
    '    End If
    'End If
    Set cnn = Nothing
    Set cnn = New ADODB.connection
    
    namaServer = lokasi_server_load(App.Path & "\data\set_db.txt")
    
    If InStr(1, namaServer, ".com", vbTextCompare) > 0 Then
        'non dsn
        cnnString = "Driver={MySQL ODBC 3.51 Driver};Server=" & namaServer & _
                ";Database=trunojoy_dbpph; User=trunojoy_user1;Password=user2017;Option=3;"
    ElseIf LCase(namaServer) = "localhost" Then
        'non dsn
        cnnString = "Driver={MySQL ODBC 3.51 Driver};Server=" & namaServer & _
                ";Database=trunojoy_dbpph; User=trunojoy_user1;Password=user2017;Option=3;"
    ElseIf IsNumeric(Replace(namaServer, ".", "")) = True Then
        'non dsn
        cnnString = "Driver={MySQL ODBC 3.51 Driver};Server=" & namaServer & _
                ";Database=trunojoy_dbpph; User=trunojoy_user1;Password=user2017;Option=3;"
    Else
        cnnString = "DSN=" & namaServer
    End If
    
    'uname: trunojoy_user1
    'pass: user2017
    'database : trunojoy_dbpph
    
    DoEvents
    cnn.ConnectionString = cnnString
    On Error GoTo er1
    cnn.Open
    DoEvents
    dbMySQL_open = True
    Exit Function
er1:
    MsgBox Err.DESCRIPTION & vbCr & "connection error", vbCritical
    dbMySQL_open = False
End Function


Function tbMDivisi_insert(kodeDivisi As String, nama_Divisi As String, ket As String)
    Dim sql As String
    
    If isDataAda("mdivisi", "kodedivisi", kodeDivisi, cnn) = True Then
        Call pesan2("Kode " & kodeDivisi & " sudah ada", , vbYellow)
        tbMDivisi_insert = False
        Exit Function
    Else
        If Trim(nama_Divisi) = "" Then
            Call pesan2("Data tidak lengkap", , vbYellow)
            tbMDivisi_insert = False
            Exit Function
        End If
        
        sql = "insert into mdivisi (kodedivisi, nama_divisi, ket) values ('" & _
            cleanStr(kodeDivisi) & "','" & cleanStr(nama_Divisi, True) & "','" & cleanStr(ket, True) & "')"
        If ExecSQL1(cnn, sql) <> 0 Then
            MsgBox "error run " & sql, vbInformation
            tbMDivisi_insert = False
        Else
            tbMDivisi_insert = True
        End If
    End If
End Function

Function tbMDivisi_Update(kodeDivisi As String, nama_Divisi As String, ket As String) As Boolean
    Dim sql As String
    
    sql = "update mdivisi set nama_divisi = '" & cleanStr(nama_Divisi, True) & _
            "', ket = '" & cleanStr(ket, True) & "' where kodedivisi = '" & cleanStr(kodeDivisi) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        MsgBox "error run " & sql, vbInformation
        tbMDivisi_Update = False
    Else
        tbMDivisi_Update = True
    End If
End Function

Function tbMDivisi_Delete(kodeDivisi As String) As Boolean
    Dim sql As String
    
    sql = "delete from mdivisi where kodedivisi = '" & cleanStr(kodeDivisi) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        MsgBox "error run " & sql, vbInformation
        tbMDivisi_Delete = False
    Else
        tbMDivisi_Delete = True
    End If
End Function

Function tbPengguna_insert(nuser As String, pwd1 As String, Level1 As String, kodeDivisi As String) As Boolean
    Dim sql As String
    
    If isDataAda("pengguna", "nuser", nuser, cnn) = True Then
        Call pesan2("Nama user " & nuser & " sudah ada", , vbYellow)
        tbPengguna_insert = False
        Exit Function
    Else
        If Trim(Level1) = "" Or Trim(kodeDivisi) = "" Then
            Call pesan2("Data tidak lengkap", , vbYellow)
            tbPengguna_insert = False
            Exit Function
        End If
    
    
        sql = "insert into pengguna (nuser, pwd1, level1, kodedivisi) values ('" & _
                cleanStr(nuser) & "','" & EncryptString(cleanStr(pwd1), "dvak2017") & "','" & Trim(Level1) & "','" & _
                Trim(kodeDivisi) & "')"
        If ExecSQL1(cnn, sql) <> 0 Then
            MsgBox "error run " & sql, vbInformation
            tbPengguna_insert = False
        Else
            tbPengguna_insert = True
        End If
    End If
End Function

Function tbPengguna_Update(nuser As String, pwd1 As String) As Boolean
    Dim sql As String
    
    sql = "update pengguna set pwd1 = '" & EncryptString(cleanStr(pwd1), "dvak2017") & _
            "' where nuser = '" & cleanStr(nuser) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        MsgBox "error run " & sql, vbInformation
        tbPengguna_Update = False
    Else
        tbPengguna_Update = True
    End If
End Function

Function tbPengguna_Delete(nuser As String) As Boolean
    Dim sql As String
    
    sql = "delete from pengguna where nuser = '" & cleanStr(nuser) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        MsgBox "error run " & sql, vbInformation
        tbPengguna_Delete = False
    Else
        tbPengguna_Delete = True
    End If
End Function

Function tbPengguna_getLevel1(nuser As String) As Integer
    Dim sql As String, t As String
    
    On Error GoTo er1
        sql = "select level1 from pengguna where nuser = '" & cleanStr(nuser) & "'"
        t = cari_data1(cnn, sql, True)
        tbPengguna_getLevel1 = CInt(t)
        Exit Function
er1:
        MsgBox Err.DESCRIPTION
End Function

Function tbPengguna_getDivisi(nuser As String) As String
    Dim sql As String, t As String
    
        sql = "select kodedivisi from pengguna where nuser = '" & cleanStr(nuser) & "'"
        t = cari_data1(cnn, sql)
        tbPengguna_getDivisi = t
End Function

Function tbPengguna_isValid_Password(nuser As String, pwd1 As String) As Integer
    'return level
    
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim pwdDb As String, Level1 As String
        
    sql = "select pwd1, level1 from pengguna where nuser = '" & cleanStr(nuser) & "'"
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        'can not open rs
        tbPengguna_isValid_Password = -2
        Exit Function
    End If
    
    If RecordCount(rs) > 0 Then
        pwdDb = cek_null(rs(0))
        Level1 = cek_null(rs(1))
        If EncryptString(cleanStr(pwd1), "dvak2017") = pwdDb Then
            tbPengguna_isValid_Password = cek_Int(Level1)
        Else
            tbPengguna_isValid_Password = 0
        End If
    Else
        'no records
        tbPengguna_isValid_Password = -3
        Exit Function
    End If
    
    
End Function

Function tbPph15_isDataAda(npwp_kpp As String, Nomor_Bukti_Potong As String, Pembetulan As String) As Boolean
    Dim sql As String, t As String
    
    sql = "select count(*) from pph15 where npwp_kpp = '" & Trim(npwp_kpp) & "' and nomor_bukti_potong = '" & Trim(Nomor_Bukti_Potong) & _
            "' and pembetulan = '" & Trim(Pembetulan) & "'"
    t = cari_data1(cnn, sql, True)
    If CInt(t) > 0 Then
        tbPph15_isDataAda = True
    Else
        tbPph15_isDataAda = False
    End If
End Function

Function tbPph15_insert(Kode_Form As String, Masa_Pajak As String, Tahun_Pajak As String, Pembetulan As String, _
                        npwp_wp As String, Nama_WP As String, Alamat_WP As String, Nomor_Bukti_Potong As String, _
                        Tanggal_Bukti_Potong As Date, negara_sumber_penghasilan As String, _
                        kode_option_penghasilan As String, Jumlah_Bruto As Currency, Tarif As String, _
                        pph_dipotong As Currency, invoice_ket As String, kode_divisi As String, npwp_kpp As String, _
                        kd_proyek As String, nott As String, nofaktur As String, email As String) As Integer
    
    '-- return
    '1: sukses insert
    '2: update
    '--
    
    Dim sql As String
    Dim return1 As Integer
    
    return1 = 0
    If tbPph15_isDataAda(npwp_kpp, Nomor_Bukti_Potong, Pembetulan) = True Then
        'data sudah ada, di hapus dulu
        If tbPph15_delete(npwp_kpp, Nomor_Bukti_Potong, Pembetulan) = True Then
            return1 = 2
        Else
            tbPph15_insert = -1
            Exit Function
        End If
    End If
    
    Nama_WP = cleanStr(Nama_WP)
    sql = "insert into pph15 (npwp_kpp, kode_form, masa_pajak, tahun_pajak, " & _
            "pembetulan, npwp_wp, nama_wp, " & _
            "alamat_wp, nomor_bukti_potong, tanggal_bukti_potong, " & _
            "negara_sumber_penghasilan, kode_option_penghasilan, jumlah_bruto, " & _
            "tarif, pph_dipotong, invoice_ket, " & _
            "kode_divisi, tgl_import, " & _
            "kd_proyek, nott, nofaktur, email) values ('" & Trim(npwp_kpp) & "', '" & _
            Trim(Kode_Form) & "','" & Trim(Masa_Pajak) & "','" & Trim(Tahun_Pajak) & "','" & _
            Trim(Pembetulan) & "','" & Trim(npwp_wp) & "','" & Trim(Nama_WP) & "','" & _
            Trim(Alamat_WP) & "','" & Trim(Nomor_Bukti_Potong) & "','" & set_tgl_perv(Tanggal_Bukti_Potong) & "','" & _
            Trim(negara_sumber_penghasilan) & "','" & Trim(kode_option_penghasilan) & "','" & Trim(Jumlah_Bruto) & "','" & _
            Trim(Tarif) & "','" & Trim(pph_dipotong) & "','" & Trim(invoice_ket) & "','" & _
            Trim(kode_divisi) & "','" & set_tgl_perv(Now) & "','" & _
            Trim(kd_proyek) & "','" & Trim(nott) & "','" & Trim(nofaktur) & "', '" & Trim(email) & "')"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbPph15_insert = -1
    Else
        If return1 = 2 Then
            tbPph15_insert = 2
        Else
            tbPph15_insert = 1
        End If
    End If
End Function


Function tbMNpwp_insert(npwp As String, nama As String, alamat As String, skaryawan As Integer) As Boolean
    Dim sql As String
    
    If isDataAda("mnpwp", "npwp", npwp, cnn) = True Then
        Call pesan2("NPWP " & npwp & " sudah ada", , vbYellow)
        tbMNpwp_insert = False
        Exit Function
    Else
        If Trim(nama) = "" Then
            Call pesan2("Data tidak lengkap", , vbYellow)
            tbMNpwp_insert = False
            Exit Function
        End If
    
    
        sql = "insert into mnpwp (npwp, nama, alamat, skaryawan) values ('" & _
                cleanStr(npwp) & "','" & cleanStr(nama) & "','" & cleanStr(alamat) & "','" & _
                Trim(skaryawan) & "')"
        If ExecSQL1(cnn, sql) <> 0 Then
            MsgBox "error run " & sql, vbInformation
            tbMNpwp_insert = False
        Else
            tbMNpwp_insert = True
        End If
    End If
End Function

Function tbMkaryawan_isDataAda(NIK As String, npwp As String, nama As String) As Boolean
                            
    If isDataAda("mkaryawan", "concat(nik, npwp, nama)", NIK & npwp & nama, cnn) = True Then
        tbMkaryawan_isDataAda = True
    Else
        tbMkaryawan_isDataAda = False
    End If
End Function

Function tbMkaryawan_isDataAda2(npwp As String, nama As String) As Boolean
                            
    If isDataAda("mkaryawan", "concat(npwp, nama)", npwp & nama, cnn) = True Then
        tbMkaryawan_isDataAda2 = True
    Else
        tbMkaryawan_isDataAda2 = False
    End If
End Function

Function tbMkaryawan_insert(NIK As String, npwp As String, nama As String, alamat As String, _
                            jenis_kelamin As String, ptkp As String) As Boolean
    Dim sql As String
    Dim klm(), isi()
    
    If isDataAda("mkaryawan", "concat(nik, npwp, nama)", NIK & npwp & nama, cnn) = True Then
        Call pesan2("NIK/NPWP/NAMA " & NIK & "/" & npwp & "/" & nama & " sudah ada", , vbYellow)
        tbMkaryawan_insert = False
        Exit Function
    Else
        If Trim(nama) = "" Then
            Call pesan2("Data tidak lengkap", , vbYellow)
            tbMkaryawan_insert = False
            Exit Function
        End If
    
        klm = Array("nik", "npwp", "nama", "alamat", "jenis_kelamin", "ptkp")
        isi = Array(NIK, npwp, nama, alamat, jenis_kelamin, ptkp)
    
        If tbInsert("mkaryawan", klm, isi, cnn) = True Then
            tbMkaryawan_insert = True
        Else
            MsgBox "error run " & sql, vbInformation
            tbMkaryawan_insert = False
        End If
    End If
End Function

Function tbAccpac_insert(tahun As String, accnum As String, desc As String, debits As Currency, _
                        credits As Currency) As Boolean
    Dim sql As String
    Dim klm(), isi()
    
    If isDataAda("tbaccpac", "concat(tahun, accnum)", tahun & accnum, cnn) = True Then
        'del dulu
        klm = Array("tahun", "accnum")
        isi = Array(tahun, accnum)
        If tbDelete("tbaccpac", klm, isi, cnn) = False Then
            tbAccpac_insert = False
            Exit Function
        End If
    End If
        
        klm = Array("tahun", "accnum", "desc1", "debits", "credits", "kdproyek")
        isi = Array(tahun, accnum, cek_null(desc), debits, credits, Right(Trim(accnum), 6))
    
        If tbInsert("tbaccpac", klm, isi, cnn) = True Then
            tbAccpac_insert = True
        Else
            MsgBox "error run " & sql, vbInformation
            tbAccpac_insert = False
        End If
End Function

Function tbMNpwp_Delete(npwp As String) As Boolean
    Dim sql As String
    
    sql = "delete from mnpwp where npwp = '" & cleanStr(npwp) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        MsgBox "error run " & sql, vbInformation
        tbMNpwp_Delete = False
    Else
        tbMNpwp_Delete = True
    End If
End Function

Function tbMNpwp_setStatus(npwp As String) As Boolean
    Dim sql As String, t As String
    
    'cek status lama
    sql = "select skaryawan from mnpwp where npwp = '" & cleanStr(npwp) & "'"
    t = cari_data1(cnn, sql, True)
    If CInt(t) = 0 Then
        t = "1"
    Else
        t = "0"
    End If
    
    sql = "update mnpwp set skaryawan = '" & t & "' where npwp = '" & cleanStr(npwp) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        MsgBox "error run " & sql, vbInformation
        tbMNpwp_setStatus = False
    Else
        tbMNpwp_setStatus = True
    End If
End Function

Function tbPph15_delete(npwp_kpp As String, Nomor_Bukti_Potong As String, Pembetulan As String) As Boolean
    Dim sql As String
    
    sql = "delete from pph15 where npwp_kpp = '" & Trim(npwp_kpp) & "' and nomor_bukti_potong = '" & Trim(Nomor_Bukti_Potong) & _
            "' and pembetulan = '" & Trim(Pembetulan) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        tbPph15_delete = False
    Else
        tbPph15_delete = True
    End If
End Function



Function tbPphX_deleteById(id1 As String, namaPPh As String) As Boolean
    Dim sql As String
    
    sql = "delete from " & namaPPh & " where id1 = '" & Trim(id1) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        tbPphX_deleteById = False
    Else
        tbPphX_deleteById = True
    End If
End Function

Function tbPphX_deleteByKPP(namaPPh As String, npwp_kpp As String, tahun As String, masa As String, DIVISI As String) As Boolean
    Dim sql As String
    Dim p
    
    If npwp_kpp = "ALL" Then
        p = MsgBox("Pilihan KPP = ALL." & vbCr & "Yakin akan menghapus data untuk semua KPP?", vbYesNo)
        If p = vbYes Then
            sql = "delete from " & namaPPh & " where Tahun_Pajak = '" & tahun & "' and Masa_Pajak = '" & masa & _
                    "' and kode_divisi = '" & DIVISI & "' "
        Else
            tbPphX_deleteByKPP = True
            Exit Function
        End If
    Else
        sql = "delete from " & namaPPh & " where Tahun_Pajak = '" & tahun & "' and Masa_Pajak = '" & masa & _
            "' and kode_divisi = '" & DIVISI & "' and NPWP_KPP = '" & npwp_kpp & "'"
    End If
    
    sql = "delete from " & namaPPh & " where Tahun_Pajak = '" & tahun & "' and Masa_Pajak = '" & masa & _
            "' and kode_divisi = '" & DIVISI & "' and NPWP_KPP = '" & npwp_kpp & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        tbPphX_deleteByKPP = False
    Else
        tbPphX_deleteByKPP = True
    End If
End Function

Function tbPphX_editById(id1 As String, namaPPh As String, nott As String, nofaktur As String, kdPROYEK As String) As Boolean
    Dim sql As String
    
    If Trim(namaPPh) = "pph15" Or Trim(namaPPh) = "pph22" Or Trim(namaPPh) = "pph23" Or _
        Trim(namaPPh) = "pph26" Or Trim(namaPPh) = "pph42_konstruksi" Or Trim(namaPPh) = "pph42_sewa" Then
    
            sql = "update " & namaPPh & " set kd_proyek = '" & Trim(kdPROYEK) & "', nott = '" & _
                    Trim(nott) & "', nofaktur = '" & Trim(nofaktur) & "' where id1 = '" & Trim(id1) & "'"
            If ExecSQL1(cnn, sql) <> 0 Then
                tbPphX_editById = False
            Else
                tbPphX_editById = True
            End If
    End If
End Function

Function tbPph23_isDataAda(Tahun_Pajak As String, npwp_kpp As String, Nomor_Bukti_Potong As String, Pembetulan As String) As Boolean
    Dim sql As String, t As String
    
    sql = "select count(*) from pph23 where NPWP_KPP = '" & Trim(npwp_kpp) & _
            "' and Nomor_Bukti_Potong = '" & Trim(Nomor_Bukti_Potong) & _
            "' and Pembetulan = '" & Trim(Pembetulan) & "' and Tahun_Pajak = '" & _
            Trim(Tahun_Pajak) & "'"
    t = cari_data1(cnn, sql, True)
    If CInt(t) > 0 Then
        tbPph23_isDataAda = True
    Else
        tbPph23_isDataAda = False
    End If
End Function

Function tbPph23_delete(Tahun_Pajak As String, npwp_kpp As String, Nomor_Bukti_Potong As String, Pembetulan As String) As Boolean
    Dim sql As String
    
    sql = "delete from pph23 where NPWP_KPP = '" & Trim(npwp_kpp) & _
            "' and Nomor_Bukti_Potong = '" & Trim(Nomor_Bukti_Potong) & _
            "' and Pembetulan = '" & Trim(Pembetulan) & "' and Tahun_Pajak = '" & _
            Trim(Tahun_Pajak) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        tbPph23_delete = False
    Else
        tbPph23_delete = True
    End If
End Function


Function tbPph23_insert(npwp_kpp As String, Kode_Form As String, Masa_Pajak As String, Tahun_Pajak As String, _
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
                    kd_proyek As String, nott As String, nofaktur As String, email As String) As Integer
    
    
    
    '-- return
    '1: sukses insert
    '2: update
    '--
    
    Dim sql As String
    Dim return1 As Integer
    
    return1 = 0
    If tbPph23_isDataAda(Tahun_Pajak, npwp_kpp, Nomor_Bukti_Potong, Pembetulan) = True Then
        'data sudah ada, di hapus dulu
        If tbPph23_delete(Tahun_Pajak, npwp_kpp, Nomor_Bukti_Potong, Pembetulan) = True Then
            return1 = 2
        Else
            tbPph23_insert = -1
            Exit Function
        End If
    End If
    
    
    'yang di skip
    'Nilai_Bruto_9 s/d Nilai_Bruto_13
    'Kode_Jasa_6d2 s/d Kode_Jasa_6d6
    
    Nama_WP = cleanStr(Nama_WP)
    
    sql = "insert into pph23(NPWP_KPP, Kode_Form, Masa_Pajak, " & _
            "Tahun_Pajak, Pembetulan, NPWP_WP, " & _
            "Nama_WP, Alamat_WP, Nomor_Bukti_Potong, " & _
            "Tanggal_Bukti_Potong, " & _
            "Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1, " & _
            "Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, " & _
            "Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, " & _
            "Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, " & _
            "Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, " & _
            "Nilai_Bruto_6a, Tarif_6a, PPh_Yang_Dipotong__6a, " & _
            "Nilai_Bruto_6b, Tarif_6b, PPh_Yang_Dipotong__6b, " & _
            "Nilai_Bruto_6c, Tarif_6c, PPh_Yang_Dipotong__6c, " & _
            "Kode_Jasa_6d1, Nilai_Bruto_6d1, Tarif_6d1, PPh_Yang_Dipotong__6d1, "
    sql = sql & _
            "Jumlah_Nilai_Bruto_, Jumlah_PPh_Yang_Dipotong, kode_divisi, tgl_import, " & _
            "kd_proyek, nott, nofaktur, email) values ('" & _
            Trim(npwp_kpp) & "', '" & Trim(Kode_Form) & "','" & Trim(Masa_Pajak) & "','" & _
            Trim(Tahun_Pajak) & "','" & Trim(Pembetulan) & "','" & Trim(npwp_wp) & "','" & _
            Trim(Nama_WP) & "','" & Trim(Alamat_WP) & "','" & Trim(Nomor_Bukti_Potong) & "','" & _
            set_tgl_perv(Tanggal_Bukti_Potong) & "','" & _
            Trim(Nilai_Bruto_1) & "','" & Trim(Tarif_1) & "','" & Trim(PPh_Yang_Dipotong__1) & "','" & _
            Trim(Nilai_Bruto_2) & "','" & Trim(Tarif_2) & "','" & Trim(PPh_Yang_Dipotong__2) & "','" & _
            Trim(Nilai_Bruto_3) & "','" & Trim(Tarif_3) & "','" & Trim(PPh_Yang_Dipotong__3) & "','" & _
            Trim(Nilai_Bruto_4) & "','" & Trim(Tarif_4) & "','" & Trim(PPh_Yang_Dipotong__4) & "','" & _
            Trim(Nilai_Bruto_5) & "','" & Trim(Tarif_5) & "','" & Trim(PPh_Yang_Dipotong__5) & "','" & _
            Trim(Nilai_Bruto_6a) & "','" & Trim(Tarif_6a) & "','" & Trim(PPh_Yang_Dipotong__6a) & "','" & _
            Trim(Nilai_Bruto_6b) & "','" & Trim(Tarif_6b) & "','" & Trim(PPh_Yang_Dipotong__6b) & "','" & _
            Trim(Nilai_Bruto_6c) & "','" & Trim(Tarif_6c) & "','" & Trim(PPh_Yang_Dipotong__6c) & "','" & _
            Trim(Kode_Jasa_6d1) & "','" & Trim(Nilai_Bruto_6d1) & "','" & Trim(Tarif_6d1) & "','" & Trim(PPh_Yang_Dipotong__6d1) & "','" & _
            Trim(Jumlah_Nilai_Bruto_) & "','" & Trim(Jumlah_PPh_Yang_Dipotong) & "','" & Trim(kode_divisi) & "','" & set_tgl_perv(Now) & "','" & _
            Trim(kd_proyek) & "','" & Trim(nott) & "','" & Trim(nofaktur) & "', '" & Trim(email) & "')"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbPph23_insert = -1
    Else
        If return1 = 2 Then
            tbPph23_insert = 2
        Else
            tbPph23_insert = 1
        End If
    End If
End Function

Function tbPph21tf_delete(npwp_kpp As String, Nomor_Bukti_Potong As String, Pembetulan As String) As Boolean
    Dim sql As String
    
    sql = "delete from pph21tf where NPWP_KPP = '" & Trim(npwp_kpp) & _
            "' and Nomor_Bukti_Potong = '" & Trim(Nomor_Bukti_Potong) & _
            "' and Pembetulan = '" & Trim(Pembetulan) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        tbPph21tf_delete = False
    Else
        tbPph21tf_delete = True
    End If
End Function

Function tbPph21tf_isDataAda(npwp_kpp As String, Nomor_Bukti_Potong As String, Pembetulan As String) As Boolean
    Dim sql As String, t As String
    
    sql = "select count(*) from pph21tf where NPWP_KPP = '" & Trim(npwp_kpp) & _
            "' and Nomor_Bukti_Potong = '" & Trim(Nomor_Bukti_Potong) & _
            "' and Pembetulan = '" & Trim(Pembetulan) & "'"
    t = cari_data1(cnn, sql, True)
    If CInt(t) > 0 Then
        tbPph21tf_isDataAda = True
    Else
        tbPph21tf_isDataAda = False
    End If
End Function

Function tbPph21pesangon_isDataAda(npwp_kpp As String, Nomor_Bukti_Potong As String, _
                            Pembetulan As String, NIK As String, nama As String) As Boolean
    Dim sql As String, t As String
    
    sql = "select count(*) from pph21pesangon where NPWP_KPP = '" & Trim(npwp_kpp) & _
            "' and Nomor_Bukti_Potong = '" & Trim(Nomor_Bukti_Potong) & _
            "' and Pembetulan = '" & Trim(Pembetulan) & "' and NIK = '" & _
            Trim(NIK) & "' and Nama = '" & Trim(nama) & "'"
    t = cari_data1(cnn, sql, True)
    If CInt(t) > 0 Then
        tbPph21pesangon_isDataAda = True
    Else
        tbPph21pesangon_isDataAda = False
    End If
End Function

Function tbPph21tf_insert(npwp_kpp As String, Masa_Pajak As String, Tahun_Pajak As String, _
                        Pembetulan As String, Nomor_Bukti_Potong As String, npwp As String, _
                        NIK As String, nama As String, alamat As String, WP_Luar_Negeri As String, _
                        Kode_Negara As String, Kode_Pajak As String, Jumlah_Bruto As Currency, _
                        Jumlah_DPP As Currency, Tanpa_NPWP As Currency, Tarif As String, _
                        Jumlah_PPh As Currency, NPWP_Pemotong As String, Nama_Pemotong As String, _
                        Tanggal_Bukti_Potong As Date, kode_divisi As String, kd_proyek As String, email As String) As Integer
    
    '-- return
    '1: sukses insert
    '2: update
    '3: skip
    '--
    
    Dim sql As String
    Dim return1 As Integer
    
    return1 = 0
    If tbPph21tf_isDataAda(npwp_kpp, Nomor_Bukti_Potong, Pembetulan) = True Then
        
        '-- update, langsung skip
        tbPph21tf_insert = 3
        Exit Function
        '---
    
        'data sudah ada, di hapus dulu
        'If tbPph21tf_delete(NPWP_KPP, Nomor_Bukti_Potong, Pembetulan) = True Then
        '    return1 = 2
        'Else
        '    tbPph21tf_insert = -1
        '    Exit Function
        'End If
    End If
    
    Nama_Pemotong = cleanStr(Nama_Pemotong)
    nama = cleanStr(nama)
    sql = "insert into pph21tf(NPWP_KPP, Masa_Pajak, Tahun_Pajak, " & _
            "Pembetulan, Nomor_Bukti_Potong, NPWP, " & _
            "NIK, Nama, Alamat, " & _
            "WP_Luar_Negeri, Kode_Negara, Kode_Pajak, " & _
            "Jumlah_Bruto, Jumlah_DPP, Tanpa_NPWP, " & _
            "Tarif, Jumlah_PPh, NPWP_Pemotong, " & _
            "Nama_Pemotong, Tanggal_Bukti_Potong, kode_divisi, tgl_import, kd_proyek, email) values ('" & _
            Trim(npwp_kpp) & "','" & Trim(Masa_Pajak) & "','" & Trim(Tahun_Pajak) & "','" & _
            Trim(Pembetulan) & "','" & Trim(Nomor_Bukti_Potong) & "','" & Trim(npwp) & "','" & _
            Trim(NIK) & "','" & Trim(nama) & "','" & Trim(alamat) & "','" & _
            Trim(WP_Luar_Negeri) & "','" & Trim(Kode_Negara) & "','" & Trim(Kode_Pajak) & "','" & _
            Trim(Jumlah_Bruto) & "','" & Trim(Jumlah_DPP) & "','" & Trim(Tanpa_NPWP) & "','" & _
            Trim(Tarif) & "','" & Trim(Jumlah_PPh) & "','" & Trim(NPWP_Pemotong) & "','" & _
            Trim(Nama_Pemotong) & "','" & set_tgl_perv(Tanggal_Bukti_Potong) & "','" & Trim(kode_divisi) & "','" & set_tgl_perv(Now) & "','" & Trim(kd_proyek) & "', '" & Trim(email) & "')"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbPph21tf_insert = -1
    Else
        If return1 = 2 Then
            tbPph21tf_insert = 2
        Else
            tbPph21tf_insert = 1
        End If
    End If
End Function

Function tbPph21pesangon_insert(npwp_kpp As String, Masa_Pajak As String, Tahun_Pajak As String, _
                        Pembetulan As String, Nomor_Bukti_Potong As String, npwp As String, _
                        NIK As String, nama As String, alamat As String, _
                        Kode_Pajak As String, Jumlah_Bruto As Currency, _
                        Tarif As String, _
                        Jumlah_PPh As Currency, NPWP_Pemotong As String, Nama_Pemotong As String, _
                        Tanggal_Bukti_Potong As Date, kode_divisi As String, kd_proyek As String, email As String) As Integer
    
    '-- return
    '1: sukses insert
    '2: update
    '3: skip
    '--
    
    Dim sql As String
    Dim return1 As Integer
    
    return1 = 0
    If tbPph21pesangon_isDataAda(npwp_kpp, Nomor_Bukti_Potong, Pembetulan, NIK, _
        nama) = True Then
        
        '-- update, langsung skip
        tbPph21pesangon_insert = 3
        Exit Function
        '---
    
        'data sudah ada, di hapus dulu
        'If tbPph21tf_delete(NPWP_KPP, Nomor_Bukti_Potong, Pembetulan) = True Then
        '    return1 = 2
        'Else
        '    tbPph21tf_insert = -1
        '    Exit Function
        'End If
    End If
    
    Nama_Pemotong = cleanStr(Nama_Pemotong)
    nama = cleanStr(nama)
    sql = "insert into pph21pesangon(NPWP_KPP, Masa_Pajak, Tahun_Pajak, " & _
            "Pembetulan, Nomor_Bukti_Potong, NPWP, " & _
            "NIK, Nama, Alamat, " & _
            "Kode_Pajak, " & _
            "Jumlah_Bruto, " & _
            "Tarif, Jumlah_PPh, NPWP_Pemotong, " & _
            "Nama_Pemotong, Tanggal_Bukti_Potong, kode_divisi, tgl_import, kd_proyek, email) values ('" & _
            Trim(npwp_kpp) & "','" & Trim(Masa_Pajak) & "','" & Trim(Tahun_Pajak) & "','" & _
            Trim(Pembetulan) & "','" & Trim(Nomor_Bukti_Potong) & "','" & Trim(npwp) & "','" & _
            Trim(NIK) & "','" & Trim(nama) & "','" & Trim(alamat) & "','" & _
            Trim(Kode_Pajak) & "','" & _
            Trim(Jumlah_Bruto) & "','" & _
            Trim(Tarif) & "','" & Trim(Jumlah_PPh) & "','" & Trim(NPWP_Pemotong) & "','" & _
            Trim(Nama_Pemotong) & "','" & set_tgl_perv(Tanggal_Bukti_Potong) & "','" & Trim(kode_divisi) & "','" & set_tgl_perv(Now) & "','" & Trim(kd_proyek) & "', '" & Trim(email) & "')"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbPph21pesangon_insert = -1
    Else
        If return1 = 2 Then
            tbPph21pesangon_insert = 2
        Else
            tbPph21pesangon_insert = 1
        End If
    End If
End Function

Function tbPph21Bulanan_delete(npwp_kpp As String, Masa_Pajak As String, Tahun_Pajak As String, Pembetulan As String, _
                                npwp As String, nama As String, email As String) As Boolean
    Dim sql As String
    
    sql = "delete from pph21bulanan where NPWP_KPP = '" & Trim(npwp_kpp) & _
            "' and Masa_Pajak = '" & Trim(Masa_Pajak) & _
            "' and Tahun_Pajak = '" & Trim(Tahun_Pajak) & _
            "' and NPWP = '" & Trim(npwp) & _
            "' and Nama = '" & Trim(nama) & _
            "' and Pembetulan = '" & Trim(Pembetulan) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        tbPph21Bulanan_delete = False
    Else
        tbPph21Bulanan_delete = True
    End If
End Function

Function tbPph21Bulanan_isDataAda(npwp_kpp As String, Masa_Pajak As String, Tahun_Pajak As String, Pembetulan As String, _
                                npwp As String, nama As String) As Boolean
    Dim sql As String, t As String
    
    sql = "select count(*) from pph21bulanan where NPWP_KPP = '" & Trim(npwp_kpp) & _
            "' and Masa_Pajak = '" & Trim(Masa_Pajak) & _
            "' and Tahun_Pajak = '" & Trim(Tahun_Pajak) & _
            "' and NPWP = '" & Trim(npwp) & _
            "' and Nama = '" & Trim(nama) & _
            "' and Pembetulan = '" & Trim(Pembetulan) & "'"
    t = cari_data1(cnn, sql, True)
    If CInt(t) > 0 Then
        tbPph21Bulanan_isDataAda = True
    Else
        tbPph21Bulanan_isDataAda = False
    End If
End Function

Function tbPph21bwhptkp_delete(npwp_kpp As String, Masa_Pajak As String, Tahun_Pajak As String, _
                                Pembetulan As String, Jumlah_karyawan As Integer, _
                                Jumlah_Bruto As Currency, kd_proyek As String) As Boolean
    Dim sql As String
    
    sql = "delete from pph21_bwhptkp where NPWP_KPP = '" & Trim(npwp_kpp) & _
            "' and Masa_Pajak = '" & Trim(Masa_Pajak) & _
            "' and Tahun_Pajak = '" & Trim(Tahun_Pajak) & _
            "' and Pembetulan = '" & Trim(Pembetulan) & _
            "' and kd_proyek = '" & Trim(kd_proyek) & _
            "' and Jumlah_karyawan = '" & Trim(Jumlah_karyawan) & _
            "' and Jumlah_Bruto = '" & Trim(Jumlah_Bruto) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        tbPph21bwhptkp_delete = False
    Else
        tbPph21bwhptkp_delete = True
    End If
End Function

Function tbPph21bwhptkp_isDataAda(npwp_kpp As String, Masa_Pajak As String, Tahun_Pajak As String, _
                                Pembetulan As String, Jumlah_karyawan As Integer, _
                                Jumlah_Bruto As Currency, kd_proyek As String) As Boolean
    Dim sql As String, t As String
    
    sql = "select count(*) from pph21_bwhptkp where NPWP_KPP = '" & Trim(npwp_kpp) & _
            "' and Masa_Pajak = '" & Trim(Masa_Pajak) & _
            "' and Tahun_Pajak = '" & Trim(Tahun_Pajak) & _
            "' and Pembetulan = '" & Trim(Pembetulan) & _
            "' and kd_proyek = '" & Trim(kd_proyek) & _
            "' and Jumlah_karyawan = '" & Trim(Jumlah_karyawan) & _
            "' and Jumlah_Bruto = '" & Trim(Jumlah_Bruto) & "'"
    t = cari_data1(cnn, sql, True)
    If CInt(t) > 0 Then
        tbPph21bwhptkp_isDataAda = True
    Else
        tbPph21bwhptkp_isDataAda = False
    End If
End Function

Function tbPph21Bulanan_insert(npwp_kpp As String, Masa_Pajak As String, Tahun_Pajak As String, _
                                Pembetulan As String, npwp As String, nama As String, Kode_Pajak As String, _
                                Jumlah_Bruto As Currency, Jumlah_PPh As Currency, Kode_Negara As String, _
                                kode_divisi As String, kd_proyek As String, email As String) As Integer
    
    '-- return
    '1: sukses insert
    '2: update
    '3: skip
    '--
    
    Dim sql As String
    Dim return1 As Integer
    
    return1 = 0
    If tbPph21Bulanan_isDataAda(npwp_kpp, Masa_Pajak, Tahun_Pajak, Pembetulan, npwp, nama) = True Then
        
        '-- update, langsung skip
        tbPph21Bulanan_insert = 3
        Exit Function
        '---
        
        'data sudah ada, di hapus dulu
        'If tbPph21Bulanan_delete(NPWP_KPP, Masa_Pajak, Tahun_Pajak, Pembetulan, npwp, nama) = True Then
        '    return1 = 2
        'Else
        '    tbPph21Bulanan_insert = -1
        '    Exit Function
        'End If
    End If
    
    nama = cleanStr(nama)
    sql = "insert into pph21bulanan(NPWP_KPP, Masa_Pajak, Tahun_Pajak, " & _
            "Pembetulan, NPWP, Nama, " & _
            "Kode_Pajak, Jumlah_Bruto, Jumlah_PPh, " & _
            "Kode_Negara, kode_divisi, tgl_import, kd_proyek, email) values ('" & _
            Trim(npwp_kpp) & "','" & Trim(Masa_Pajak) & "','" & Trim(Tahun_Pajak) & "','" & _
            Trim(Pembetulan) & "','" & Trim(npwp) & "','" & Trim(nama) & "','" & _
            Trim(Kode_Pajak) & "','" & Trim(Jumlah_Bruto) & "','" & Trim(Jumlah_PPh) & "','" & _
            Trim(Kode_Negara) & "','" & Trim(kode_divisi) & "','" & set_tgl_perv(Now) & "','" & Trim(kd_proyek) & "', '" & Trim(email) & "')"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbPph21Bulanan_insert = -1
    Else
        If return1 = 2 Then
            tbPph21Bulanan_insert = 2
        Else
            tbPph21Bulanan_insert = 1
        End If
    End If
End Function

Function tbPph21bwhptkp_insert(npwp_kpp As String, Masa_Pajak As String, Tahun_Pajak As String, _
                                Pembetulan As String, Jumlah_karyawan As Integer, _
                                Jumlah_Bruto As Currency, kode_divisi As String, _
                                kd_proyek As String) As Integer
    
    '-- return
    '1: sukses insert
    '2: update
    '3: skip
    '--
    
    Dim sql As String
    Dim return1 As Integer
    
    return1 = 0
    If tbPph21bwhptkp_isDataAda(npwp_kpp, Masa_Pajak, Tahun_Pajak, Pembetulan, _
                                Jumlah_karyawan, Jumlah_Bruto, kd_proyek) = True Then
        
        '-- update, langsung skip
        'tbPph21bwhptkp_insert = 3
        'Exit Function
        '---
        
        'data sudah ada, di hapus dulu
        If tbPph21bwhptkp_delete(npwp_kpp, Masa_Pajak, Tahun_Pajak, Pembetulan, _
                                Jumlah_karyawan, Jumlah_Bruto, kd_proyek) = True Then
            return1 = 2
        Else
            tbPph21bwhptkp_insert = -1
            Exit Function
        End If
    End If
    
    sql = "insert into pph21_bwhptkp(NPWP_KPP, Masa_Pajak, Tahun_Pajak, " & _
            "Pembetulan, Jumlah_karyawan, Jumlah_Bruto, " & _
            "kode_divisi, tgl_import, kd_proyek) values ('" & _
            Trim(npwp_kpp) & "','" & Trim(Masa_Pajak) & "','" & Trim(Tahun_Pajak) & "','" & _
            Trim(Pembetulan) & "','" & Trim(Jumlah_karyawan) & "','" & Trim(Jumlah_Bruto) & "','" & _
            Trim(kode_divisi) & "','" & set_tgl_perv(Now) & "','" & Trim(kd_proyek) & "')"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbPph21bwhptkp_insert = -1
    Else
        If return1 = 2 Then
            tbPph21bwhptkp_insert = 2
        Else
            tbPph21bwhptkp_insert = 1
        End If
    End If
End Function

Function tbPphX_delete(Tahun_Pajak As String, npwp_kpp As String, Nomor_Bukti_Potong As String, Pembetulan As String, nmTable1 As String) As Boolean
    Dim sql As String
    
    sql = "delete from " & nmTable1 & " where NPWP_KPP = '" & Trim(npwp_kpp) & _
            "' and Nomor_Bukti_Potong = '" & Trim(Nomor_Bukti_Potong) & _
            "' and Pembetulan = '" & Trim(Pembetulan) & "' and Tahun_Pajak = '" & _
            Trim(Tahun_Pajak) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        tbPphX_delete = False
    Else
        tbPphX_delete = True
    End If
End Function

Function tbPphX_isDataAda(Tahun_Pajak As String, npwp_kpp As String, Nomor_Bukti_Potong As String, Pembetulan As String, nmTable1 As String) As Boolean
    Dim sql As String, t As String
    
    sql = "select count(*) from " & nmTable1 & " where NPWP_KPP = '" & Trim(npwp_kpp) & _
            "' and Nomor_Bukti_Potong = '" & Trim(Nomor_Bukti_Potong) & _
            "' and Pembetulan = '" & Trim(Pembetulan) & "' and Tahun_Pajak = '" & _
            Trim(Tahun_Pajak) & "'"
    t = cari_data1(cnn, sql, True)
    If CInt(t) > 0 Then
        tbPphX_isDataAda = True
    Else
        tbPphX_isDataAda = False
    End If
End Function


Function tbPph21Tahunan_insert(npwp_kpp As String, Masa_Pajak As String, Tahun_Pajak As String, _
                            Pembetulan As String, Nomor_Bukti_Potong As String, Masa_Perolehan_Awal As String, _
                            Masa_Perolehan_Akhir As String, npwp As String, NIK As String, nama As String, _
                            alamat As String, jenis_kelamin As String, Status_PTKP As String, _
                            Jumlah_Tanggungan As String, Nama_Jabatan As String, WP_Luar_Negeri As String, _
                            Kode_Negara As String, Kode_Pajak As String, Jumlah_1 As Currency, _
                            Jumlah_2 As Currency, Jumlah_3 As Currency, Jumlah_4 As Currency, _
                            Jumlah_5 As Currency, Jumlah_6 As Currency, Jumlah_7 As Currency, _
                            Jumlah_8 As Currency, Jumlah_9 As Currency, Jumlah_10 As Currency, _
                            Jumlah_11 As Currency, Jumlah_12 As Currency, Jumlah_13 As Currency, _
                            Jumlah_14 As Currency, Jumlah_15 As Currency, Jumlah_16 As Currency, _
                            Jumlah_17 As Currency, Jumlah_18 As Currency, Jumlah_19 As Currency, _
                            Jumlah_20 As Currency, Status_Pindah As String, NPWP_Pemotong As String, _
                            Nama_Pemotong As String, Tanggal_Bukti_Potong As Date, kode_divisi As String, _
                            kd_proyek As String, email As String) As Integer
    
    '-- return
    '1: sukses insert
    '2: update
    '3: skip
    '--
    
    Dim sql As String
    Dim return1 As Integer
    
    return1 = 0
    If tbPphX_isDataAda(Tahun_Pajak, npwp_kpp, Nomor_Bukti_Potong, Pembetulan, "pph21tahunan") = True Then
        
        '-- update, langsung skip
        tbPph21Tahunan_insert = 3
        Exit Function
        '---
        
        'data sudah ada, di hapus dulu
        'If tbPphX_delete(NPWP_KPP, Nomor_Bukti_Potong, Pembetulan, "pph21tahunan") = True Then
        '    return1 = 2
        'Else
        '    tbPph21Tahunan_insert = -1
        '    Exit Function
        'End If
    End If
    
    nama = cleanStr(nama)
    sql = "insert into pph21tahunan(NPWP_KPP, Masa_Pajak, Tahun_Pajak, " & _
            "Pembetulan, Nomor_Bukti_Potong, Masa_Perolehan_Awal, " & _
            "Masa_Perolehan_Akhir, NPWP, NIK, " & _
            "Nama, Alamat, Jenis_Kelamin, " & _
            "Status_PTKP, Jumlah_Tanggungan, Nama_Jabatan, " & _
            "WP_Luar_Negeri, Kode_Negara, Kode_Pajak, " & _
            "Jumlah_1,  Jumlah_2, Jumlah_3, " & _
            "Jumlah_4, Jumlah_5, Jumlah_6, " & _
            "Jumlah_7, Jumlah_8, Jumlah_9, " & _
            "Jumlah_10, Jumlah_11, Jumlah_12, " & _
            "Jumlah_13, Jumlah_14, Jumlah_15, " & _
            "Jumlah_16, Jumlah_17, Jumlah_18, " & _
            "Jumlah_19, Jumlah_20, Status_Pindah, " & _
            "NPWP_Pemotong, Nama_Pemotong, Tanggal_Bukti_Potong, " & _
            "kode_divisi, tgl_import, kd_proyek, email) values ('" & _
            Trim(npwp_kpp) & "','" & Trim(Masa_Pajak) & "','" & Trim(Tahun_Pajak) & "','" & _
            Trim(Pembetulan) & "','" & Trim(Nomor_Bukti_Potong) & "','" & Trim(Masa_Perolehan_Awal) & "','" & _
            Trim(Masa_Perolehan_Akhir) & "','" & Trim(npwp) & "','" & Trim(NIK) & "','" & _
            Trim(nama) & "','" & Trim(alamat) & "','" & Trim(jenis_kelamin) & "','" & _
            Trim(Status_PTKP) & "','" & Trim(Jumlah_Tanggungan) & "','" & Trim(Nama_Jabatan) & "','" & _
            Trim(WP_Luar_Negeri) & "','" & Trim(Kode_Negara) & "','" & Trim(Kode_Pajak) & "','" & _
            Trim(Jumlah_1) & "','" & Trim(Jumlah_2) & "','" & Trim(Jumlah_3) & "','" & _
            Trim(Jumlah_4) & "','" & Trim(Jumlah_5) & "','" & Trim(Jumlah_6) & "','" & _
            Trim(Jumlah_7) & "','" & Trim(Jumlah_8) & "','" & Trim(Jumlah_9) & "','" & _
            Trim(Jumlah_10) & "','" & Trim(Jumlah_11) & "','" & Trim(Jumlah_12) & "','"
    sql = sql & _
            Trim(Jumlah_13) & "','" & Trim(Jumlah_14) & "','" & Trim(Jumlah_15) & "','" & _
            Trim(Jumlah_16) & "','" & Trim(Jumlah_17) & "','" & Trim(Jumlah_18) & "','" & _
            Trim(Jumlah_19) & "','" & Trim(Jumlah_20) & "','" & Trim(Status_Pindah) & "','" & _
            Trim(NPWP_Pemotong) & "','" & Trim(Nama_Pemotong) & "','" & set_tgl_perv(Tanggal_Bukti_Potong) & "','" & _
            Trim(kode_divisi) & "','" & set_tgl_perv(Now) & "','" & Trim(kd_proyek) & "', '" & Trim(email) & "')"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbPph21Tahunan_insert = -1
    Else
        If return1 = 2 Then
            tbPph21Tahunan_insert = 2
        Else
            tbPph21Tahunan_insert = 1
        End If
    End If
End Function

Function tbPph21Tahunan2_isDataAda(npwp As String, NIK As String, nama As String, tahun As String, _
                                    bulan As String, npwp_kpp As String) As Boolean
    Dim sql As String, t As String
    
    sql = "select count(*) from pph21tahunan2 where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "' and Bulan = '" & _
            Trim(bulan) & "' and NPWP_KPP = '" & Trim(npwp_kpp) & "'"
    t = cari_data1(cnn, sql, True)
    If CInt(t) > 0 Then
        tbPph21Tahunan2_isDataAda = True
    Else
        tbPph21Tahunan2_isDataAda = False
    End If
End Function

Function tbPph21Tahunan2_Update2(npwp As String, NIK As String, nama As String, tahun As String, _
                                    bulan As String, npwp_kpp As String, penghasilan_netto_sblmnya As Currency, _
                                    pph21_terutang_sblmnya As Currency, nrp As String) As Boolean
    
    Dim klm(), isi()
    
    klm = Array("penghasilan_netto_sblmnya", "pph21_terutang_sblmnya", "nrp")
    isi = Array(penghasilan_netto_sblmnya, pph21_terutang_sblmnya, nrp)
    
    
    If tbUpdate("pph21tahunan2", klm, isi, cnn, _
                "NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "' and Bulan = '" & _
            Trim(bulan) & "' and NPWP_KPP = '" & Trim(npwp_kpp) & "'") = True Then
        tbPph21Tahunan2_Update2 = True
    Else
        tbPph21Tahunan2_Update2 = False
    End If
    
End Function

Function tbPph21Tahunan2_getTotal(KolomTotal As String, npwp As String, NIK As String, nama As String, _
                                tahun As String, npwp_kpp As String) As Currency

    Dim sql As String, t As String
    
    sql = "select sum(" & KolomTotal & ") from pph21tahunan2 " & _
            "where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "' and NPWP_KPP = '" & _
            Trim(npwp_kpp) & "'"
    t = cari_data1(cnn, sql, True)
    tbPph21Tahunan2_getTotal = cek_Money(t)
End Function

Sub tbPph21Tahunan2_getTotal2(npwp As String, NIK As String, nama As String, _
                                tahun As String, npwp_kpp As String, ByRef Gaji As Currency, _
                                ByRef Tnj_PPh As Currency, ByRef Tunjangan_Lain As Currency, _
                                ByRef JHT_JPN As Currency, ByRef Insentif_THR_Lainnya As Currency, _
                                ByRef Pensiun_Potongan_Lain As Currency, ByRef penghasilan_netto_sblmnya As Currency, _
                                ByRef pph21_terutang_sblmnya As Currency)

    Dim sql As String, rs As ADODB.Recordset
    
    sql = "select sum(Gaji), sum(Tnj_PPh), sum(Tunjangan_Lain), sum(JHT_JPN), " & _
            "sum(Insentif + THR + Lainnya), sum(Pensiun_Potongan_Lain), sum(penghasilan_netto_sblmnya), " & _
            "sum(pph21_terutang_sblmnya) " & _
            "from pph21tahunan2 " & _
            "where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "' and NPWP_KPP = '" & _
            Trim(npwp_kpp) & "'"
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Exit Sub
    End If
    
    If RecordCount(rs) > 0 Then
        Gaji = cek_Money(rs(0))
        Tnj_PPh = cek_Money(rs(1))
        Tunjangan_Lain = cek_Money(rs(2))
        JHT_JPN = cek_Money(rs(3))
        Insentif_THR_Lainnya = cek_Money(rs(4))
        Pensiun_Potongan_Lain = cek_Money(rs(5))
        penghasilan_netto_sblmnya = cek_Money(rs(6))
        pph21_terutang_sblmnya = cek_Money(rs(7))
    Else
        Gaji = 0
        Tnj_PPh = 0
        Tunjangan_Lain = 0
        JHT_JPN = 0
        Insentif_THR_Lainnya = 0
        Pensiun_Potongan_Lain = 0
        penghasilan_netto_sblmnya = 0
        pph21_terutang_sblmnya = 0
    End If
End Sub

Function tbPph21Tahunan2_getTotalBruto(npwp As String, NIK As String, nama As String, tahun As String) As Currency
    'gabungan  Gaji, Tnj PPh Gaji,     JHT & JPN,  Tunjangan Lain
    'per npwp dalam 1 tahun
    
    Dim sql As String, t As String
    
    sql = "select sum(Gaji + Tnj_PPh + JHT_JPN + Tunjangan_Lain + Insentif) from pph21tahunan2 " & _
            "where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "'"
    t = cari_data1(cnn, sql, True)
    tbPph21Tahunan2_getTotalBruto = cek_Money(t)
End Function

Function tbPph21Tahunan2_getBulanAkhir(npwp As String, NIK As String, nama As String, tahun As String, _
                                        npwp_kpp As String) As String
                                        
    'dari data yang ada, utk id tsb, pada kpp tsb, data bulan paling akhir bulan apa ??
    
    Dim sql As String, t As String
    
    sql = "select max(convert(Bulan, signed)) from pph21tahunan2 where NPWP = '" & Trim(npwp) & _
            "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "' and NPWP_KPP = '" & _
            Trim(npwp_kpp) & "'"
    t = cari_data1(cnn, sql, True)
    tbPph21Tahunan2_getBulanAkhir = t
    
End Function

Function tbPph21Tahunan2_getNomorBukti(npwp As String, NIK As String, nama As String, tahun As String, _
                                        npwp_kpp As String) As String
                                        
    
    
    Dim sql As String, t As String
    
    sql = "select no_urut_buktipotong from pph21tahunan2 where NPWP = '" & Trim(npwp) & _
            "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "' and NPWP_KPP = '" & _
            Trim(npwp_kpp) & "'"
    t = cari_data1(cnn, sql, True)
    tbPph21Tahunan2_getNomorBukti = t
    
End Function

Function tbPph21Tahunan2_getBulanAwal(npwp As String, NIK As String, nama As String, tahun As String, _
                                        npwp_kpp As String) As String
                                        
    'dari data yang ada, utk id tsb, pada kpp tsb, data bulan paling akhir bulan apa ??
    
    Dim sql As String, t As String
    
    sql = "select min(convert(Bulan, signed)) from pph21tahunan2 where NPWP = '" & Trim(npwp) & _
            "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "' and NPWP_KPP = '" & _
            Trim(npwp_kpp) & "'"
    t = cari_data1(cnn, sql, True)
    tbPph21Tahunan2_getBulanAwal = t
    
End Function

Function tbPph21Tahunan2_getData_byId(id1 As String, namaKolom As String) As String
    Dim sql As String, t As String
    
    sql = "select " & namaKolom & " from pph21tahunan2 where id1 = '" & id1 & "'"
    t = cari_data1(cnn, sql)
    tbPph21Tahunan2_getData_byId = Trim(t)
End Function

Sub tbPph21Tahunan2_getData_byId2(id1 As String, ByRef no_urut_bukti_potong As String, _
                                ByRef alamat As String, ByRef P_L As String, ByRef ptkp As String)
    
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    sql = "select no_urut_buktipotong, alamat, p_L, PTKP from pph21tahunan2 where id1 = '" & id1 & "'"
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Exit Sub
    End If
    
    If RecordCount(rs) > 0 Then
        no_urut_bukti_potong = cek_null(rs(0))
        alamat = cek_null(rs(1))
        P_L = cek_null(rs(2))
        ptkp = cek_null(rs(3))
    Else
        no_urut_bukti_potong = ""
        alamat = ""
        P_L = ""
        ptkp = ""
    End If
    
End Sub

Sub tbPph21Tahunan2_setBiayaJabatanPerBulan()
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim jRec As Long, c As Long
    Dim Bruto As Currency, biayaJabatan As Currency
    Dim id1, totalN As Currency
    
    sql = "update pph21tahunan2 set NPWP = '000000000000000' where NPWP = '' or NPWP is null or NPWP = '0'"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Exit Sub
    End If
    
    sql = "update pph21tahunan2 set NIK = '0000000000000000' where NIK = '' or NIK is null or NIK = '0' "
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("sql error", "", sql)
        Exit Sub
    End If
    
    '-- 5% dari bruto
    '-- bruto = gaji + tnj_pph + tnjLain + tnj_jht_jpn
    '--- jika kurang dari 500rb, + 5% * (insentif + THR + Lainnya)
    '------ tetap maksimal 500
    
    sql = "select (Gaji + Tnj_PPh + JHT_JPN + Tunjangan_Lain), id1 , (insentif + THR + lainnya) as n " & _
            "from pph21tahunan2 " & _
            "where (biaya_jabatan = 0 or biaya_jabatan is null) and " & _
            "((Gaji + Tnj_PPh + JHT_JPN + Tunjangan_Lain) > 0)"
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error sql", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rs)
    If jRec <= 0 Then Exit Sub
    
    rs.MoveFirst
    c = 1
    Do While rs.EOF = False
        Call info(2, "cek biaya_jabatan. Run " & c & "/" & jRec & ". " & Round((c / jRec) * 100, 2) & "%", _
                    frMenu1.StatusBar1)
        Bruto = cek_Money(rs(0))
        id1 = cek_null(rs(1))
        totalN = cek_Money(rs(2))
        
        biayaJabatan = 0.05 * Bruto
        If biayaJabatan < 500000 Then
            biayaJabatan = biayaJabatan + (0.05 * totalN)
            If biayaJabatan > 500000 Then biayaJabatan = 500000
        Else
            biayaJabatan = 500000
        End If
        
        'update
        sql = "update pph21tahunan2 set biaya_jabatan = '" & biayaJabatan & _
                "' where id1 = '" & id1 & "'"
        If ExecSQL1(cnn, sql) <> 0 Then
            sql = InputBox("error sql", "", sql)
            Exit Do
        End If
        
        rs.MoveNext
        c = c + 1
    Loop
    
End Sub



Function tbPph21Tahunan2_getTotalBiayaJabatan(npwp As String, NIK As String, nama As String, _
                                                tahun As String, npwp_kpp As String) As Currency
    'per npwp dalam 1 tahun
    '-- biaya jabatan : =IF((bruto*0,05)<500000;(J5*0,05);500000)
    'jadi harus di cari dulu nilai bruto untuk setiap bulannya...
    
    Dim sql As String, t As String, jmlBulan As String
    Dim totalBiayaJabatan As Currency, Tambahan As Currency, Bruto As Currency
    Dim rs As ADODB.Recordset
    
    
    
    'cek, jika jumlah data ada 12, maka ==
    'totalBiayaJabatan = Gaji + Tnj_PPh + JHT_JPN + Tunjangan_Lain + Insentif + THR + Lainnya
    sql = "select count(*), " & _
            "sum(Gaji + Tnj_PPh + JHT_JPN + Tunjangan_Lain + Insentif + THR + Lainnya) as bruto1, " & _
            "sum(Gaji + Tnj_PPh + JHT_JPN + Tunjangan_Lain ) as tambahan " & _
            "from pph21tahunan2 " & _
            "where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & _
            "' and NPWP_KPP = '" & Trim(npwp_kpp) & "'"
    'koreksi 15jan2021
    sql = "select count(*), " & _
            "sum(Gaji + Tnj_PPh + JHT_JPN + Tunjangan_Lain ) as bruto1, " & _
            "sum(Insentif + THR + Lainnya ) as tambahan " & _
            "from pph21tahunan2 " & _
            "where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & _
            "' and NPWP_KPP = '" & Trim(npwp_kpp) & "'"
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error sql", "", sql)
        tbPph21Tahunan2_getTotalBiayaJabatan = 0
        Exit Function
    Else
        If RecordCount(rs) > 0 Then
            jmlBulan = cek_null(rs(0))
            Bruto = cek_Money(cek_null(rs(1)))
            Tambahan = cek_Money(cek_null(rs(2)))
            
            If CInt(jmlBulan) = 12 Then
                totalBiayaJabatan = Bruto * 0.05
                If totalBiayaJabatan > 6000000 Then totalBiayaJabatan = 6000000
            Else
                totalBiayaJabatan = Bruto * 0.05
                If totalBiayaJabatan < (CInt(jmlBulan) * 500000) Then
                    'hitung tambahan
                    totalBiayaJabatan = totalBiayaJabatan + (CCur(Tambahan) * 0.05)
                    If totalBiayaJabatan > (CInt(jmlBulan) * 500000) Then
                        totalBiayaJabatan = (CInt(jmlBulan) * 500000)
                    End If
                Else
                    totalBiayaJabatan = (CInt(jmlBulan) * 500000)
                End If
            End If
        End If
    
        totalBiayaJabatan = Round(totalBiayaJabatan, 0)
        tbPph21Tahunan2_getTotalBiayaJabatan = totalBiayaJabatan
            
    End If
    
End Function

Function tbPph21Tahunan2_getTotalIuranPensiun(npwp As String, NIK As String, nama As String, tahun As String) As Currency
    'per npwp dalam 1 tahun
    
    Dim sql As String, t As String
    
    sql = "select sum(Pensiun_Potongan_Lain) from pph21tahunan2 " & _
            "where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "'"
    t = cari_data1(cnn, sql, True)
    tbPph21Tahunan2_getTotalIuranPensiun = cek_Money(t)
End Function



Function tbPph21Tahunan2_getTotalNetto(ByRef totalBruto As Currency, ByRef totalBiayaJabatan As Currency, _
                                        ByRef totalIuranPensiun As Currency, npwp As String, NIK As String, nama As String, tahun As String) As Currency
    ' "totalBruto -  totalBiaya Jabatan  - totalIuran Pensiun/Potongan Lain
    Dim total_netto As Currency
    Dim sql As String, t As String
    Dim rs As ADODB.Recordset
    
    'totalbruto = sum(Gaji + Tnj_PPh + JHT_JPN + Tunjangan_Lain + Insentif)
    'totalBiayaJabatan = sum(biaya_jabatan) from pph21tahunan2 " & _
    'totalIuranPensiun = sum(Pensiun_Potongan_Lain) from pph21tahunan2 " &

    'total_netto = tbPph21Tahunan2_getTotalBruto(NPWP, NIK, Nama, Tahun) - _
    '                tbPph21Tahunan2_getTotalBiayaJabatan(NPWP, NIK, Nama, Tahun) - _
    '                tbPph21Tahunan2_getTotalIuranPensiun(NPWP, NIK, Nama, Tahun)
    
    sql = "select '' as totak_netto, " & _
            "sum(Gaji + Tnj_PPh + JHT_JPN + Tunjangan_Lain + Insentif + THR ) as totalBruto, " & _
            "sum(biaya_jabatan) as totalBiayaJabatan, sum(Pensiun_Potongan_Lain) as totalIuranPensiun " & _
            "from pph21tahunan2 " & _
            "where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "'"
    'sql = InputBox("sql", "", sql)
    If OpenRecordSet(cnn, rs, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error sql", "", sql)
        totalBruto = 0
        totalBiayaJabatan = 0
        totalIuranPensiun = 0
        total_netto = 0
    End If
    
    If RecordCount(rs) <= 0 Then
        totalBruto = 0
        totalBiayaJabatan = 0
        totalIuranPensiun = 0
        total_netto = 0
    End If
    
    totalBruto = cek_Money(rs(1))
    totalBiayaJabatan = cek_Money(rs(2))
    'cek, jika jumlah data ada 12, maka ==
    'totalBiayaJabatan = Gaji + Tnj_PPh + JHT_JPN + Tunjangan_Lain + Insentif + THR + Lainnya
    sql = "select count(*) from pph21tahunan2 " & _
            "where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "'"
    t = cari_data1(cnn, sql, True)
    If CInt(t) = 12 Then
        sql = "select sum(Gaji + Tnj_PPh + JHT_JPN + Tunjangan_Lain + Insentif + THR + Lainnya) " & _
                "from pph21tahunan2 " & _
                "where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
                "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "'"
        t = cari_data1(cnn, sql, True)
        totalBiayaJabatan = Round(CCur(t) * 0.05, 0)
        If totalBiayaJabatan > 12000000 Then totalBiayaJabatan = 12000000
    End If
    
    totalIuranPensiun = cek_Money(rs(3))
    total_netto = totalBruto - totalBiayaJabatan - totalIuranPensiun
    
    tbPph21Tahunan2_getTotalNetto = total_netto
End Function

Function tbPph21Tahunan2_Delete(id1 As String) As Boolean
    'delete byId
    Dim sql As String, t As String
    
    sql = "delete from pph21tahunan2 where id1 = '" & id1 & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        tbPph21Tahunan2_Delete = False
    Else
        tbPph21Tahunan2_Delete = True
    End If
End Function

Function tbPph21Tahunan2_Delete1Kpp(kdCENTER As String, tahun As String, bulan As String, _
                                    npwp_kpp As String) As Boolean
    Dim sql As String, t As String
    
    sql = "delete from pph21tahunan2 where kdCENTER = '" & Trim(kdCENTER) & "' and Tahun = '" & _
            Trim(tahun) & "' and Bulan = '" & Trim(bulan) & "' and NPWP_KPP = '" & Trim(npwp_kpp) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        tbPph21Tahunan2_Delete1Kpp = False
    Else
        tbPph21Tahunan2_Delete1Kpp = True
    End If
End Function

Function tbPph21Tahunan2_Delete1Divisi(kdCENTER As String, tahun As String, bulan As String) As Boolean
    Dim sql As String, t As String
    
    sql = "delete from pph21tahunan2 where kdCENTER = '" & Trim(kdCENTER) & "' and Tahun = '" & _
            Trim(tahun) & "' and Bulan = '" & Trim(bulan) & "' "
    If ExecSQL1(cnn, sql) <> 0 Then
        tbPph21Tahunan2_Delete1Divisi = False
    Else
        tbPph21Tahunan2_Delete1Divisi = True
    End If
End Function

Function tbPph21Tahunan2_get_kdCENTER(npwp As String, NIK As String, nama As String, tahun As String, _
                                        npwp_kpp As String, bulanAwal As String) As String
    Dim sql As String, t As String
    
    sql = "select kdCENTER from pph21tahunan2 " & _
            "where npwp = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & "' and  nama = '" & _
            Trim(nama) & "' and Tahun = '" & Trim(tahun) & "' and " & _
            "npwp_kpp = '" & Trim(npwp_kpp) & "' and Bulan = '" & Trim(bulanAwal) & "'"
    t = cari_data1(cnn, sql, True)
    tbPph21Tahunan2_get_kdCENTER = t
End Function

Function tbPph21Tahunan2_get_ID1(npwp As String, NIK As String, nama As String, tahun As String) As String
    Dim sql As String, t As String
    
    sql = "select id1 from pph21tahunan2 " & _
            "where npwp = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & "' and  nama = '" & _
            Trim(nama) & "' and Tahun = '" & Trim(tahun) & "' "
    t = cari_data1(cnn, sql, True)
    tbPph21Tahunan2_get_ID1 = t
End Function




Function tbPph21Tahunan2_get_NilaiPTKP(id1 As String) As Currency
    Dim ptkp As String
    Dim nilaiPtkp As Currency
        
    ptkp = tbPph21Tahunan2_getData_byId(id1, "PTKP")
    nilaiPtkp = tbM_Ptkp_getNilai(ptkp)
    tbPph21Tahunan2_get_NilaiPTKP = nilaiPtkp
End Function

Function tbPph21Tahunan2_getPtkp(npwp As String, NIK As String, nama As String, tahun As String, bulan As String, ptkp) As String
    'cek data ptkp sebelumnya, jika kosong, pakai data ptkp yang sudah ada...
    'jika data ptkp yang ada masanya lebih awal, ambil data yang ada sebelumnya..
    
    Dim sql As String, t As String
    
    sql = "select PTKP From pph21tahunan2 where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "' order by cast(Bulan as int)"
    t = cari_data1(cnn, sql)
    If Trim(t) = "" Then
        tbPph21Tahunan2_getPtkp = UCase(ptkp)
    Else
        tbPph21Tahunan2_getPtkp = UCase(t)
    End If
End Function

Function tbPph21Tahunan2_getJabatan(npwp As String, NIK As String, nama As String, tahun As String) As String
    
    Dim sql As String, t As String
    
    sql = "select Jabatan From pph21tahunan2 where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "' order by cast(Bulan as int) desc "
    t = cari_data1(cnn, sql)
    tbPph21Tahunan2_getJabatan = t
End Function

Function tbPph21Tahunan2_getJmlData(npwp As String, NIK As String, nama As String, tahun As String, _
                                    npwpKpp As String) As String
    
    Dim sql As String, t As String
    
    If Trim(npwpKpp) = "" Or Trim(npwpKpp) = "ALL" Then
        sql = "select count(*) From pph21tahunan2 where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & "'"
    Else
        sql = "select count(*) From pph21tahunan2 where NPWP = '" & Trim(npwp) & "' and NIK = '" & Trim(NIK) & _
            "' and Nama = '" & Trim(nama) & "' and Tahun = '" & Trim(tahun) & _
            "' and NPWP_KPP = '" & Trim(npwpKpp) & "'"
    End If
    t = cari_data1(cnn, sql)
    tbPph21Tahunan2_getJmlData = t
End Function

Function tbPph21Tahunan2_Edit(id1 As String, nama As String, npwp As String, _
                                NIK As String, alamat As String, Jabatan As String, P_L As String, _
                                bulan As String, penghasilan_netto_sblmnya As Currency, _
                                pph21_terutang_sblmnya As Currency) As Boolean
                                
    Dim sql As String
    
    sql = "update pph21tahunan2 set Nama = '" & UCase(CekPetik(nama)) & "', NPWP = '" & UCase(CekPetik(npwp)) & _
            "', NIK = '" & UCase(CekPetik(NIK)) & "', Alamat = '" & UCase(CekPetik(alamat)) & _
            "', Jabatan = '" & UCase(CekPetik(Jabatan)) & "', P_L = '" & P_L & _
            "', Bulan = '" & Trim(bulan) & "', penghasilan_netto_sblmnya = '" & penghasilan_netto_sblmnya & _
            "', pph21_terutang_sblmnya = '" & pph21_terutang_sblmnya & "' where id1 = '" & UCase(CekPetik(id1)) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbPph21Tahunan2_Edit = False
    Else
        tbPph21Tahunan2_Edit = True
    End If

End Function

Function tbPph21Tahunan2_insert(NO1 As Long, bulan As String, tahun As String, npwp_kpp As String, _
                                kdPROYEK As String, kdCENTER As String, nama As String, npwp As String, _
                                NIK As String, alamat As String, Jabatan As String, P_L As String, _
                                ptkp As String, Gaji As Currency, Tnj_PPh As Currency, Tunjangan_Lain As Currency, _
                                JHT_JPN As Currency, Bruto As Currency, Insentif As Currency, Thr As Currency, _
                                Lainnya As Currency, Pensiun_Potongan_Lain As Currency) As Integer
    
    '-- return
    '1: sukses insert
    '3: skip
    '--
    
    Dim sql As String, t As String
    Dim skip1 As Boolean
    Dim return1 As Integer
    
    return1 = 0
    skip1 = False
    If tbPph21Tahunan2_isDataAda(npwp, NIK, nama, tahun, bulan, npwp_kpp) = True Then
        skip1 = True
    End If
    
    If skip1 = False Then
        nama = cleanStr(nama)
        t = tbPph21Tahunan2_getPtkp(npwp, NIK, nama, tahun, bulan, ptkp)
        ptkp = t
    
        sql = "insert into pph21tahunan2(No1, Bulan, Tahun, " & _
            "NPWP_KPP, kdPROYEK, kdCENTER, " & _
            "Nama, NPWP, NIK, " & _
            "Alamat, Jabatan, P_L, " & _
            "PTKP, Gaji, Tnj_PPh, " & _
            "Tunjangan_Lain, JHT_JPN, Bruto, " & _
            "Insentif, THR, Lainnya, " & _
            "Pensiun_Potongan_Lain, tglupdate) values ('" & _
            Trim(NO1) & "','" & Trim(bulan) & "','" & Trim(tahun) & "','" & _
            Trim(npwp_kpp) & "','" & Trim(kdPROYEK) & "','" & Trim(kdCENTER) & "','" & _
            Trim(nama) & "','" & Trim(npwp) & "','" & Trim(NIK) & "','" & _
            Trim(alamat) & "','" & Trim(Jabatan) & "','" & Trim(P_L) & "','" & _
            Trim(ptkp) & "','" & Trim(Gaji) & "','" & Trim(Tnj_PPh) & "','" & _
            Trim(Tunjangan_Lain) & "','" & Trim(JHT_JPN) & "','" & Trim(Bruto) & "','" & _
            Trim(Insentif) & "','" & Trim(Thr) & "','" & Trim(Lainnya) & "','" & _
            Trim(Pensiun_Potongan_Lain) & "','" & set_tgl_perv(Now) & "')"
        If ExecSQL1(cnn, sql) <> 0 Then
            sql = InputBox("", "", sql)
            tbPph21Tahunan2_insert = -1
        Else
            If return1 = 2 Then
                tbPph21Tahunan2_insert = 2
            Else
                tbPph21Tahunan2_insert = 1
            End If
        End If
    Else
        tbPph21Tahunan2_insert = 3
    End If
End Function


Function tbPph21Tahunan2_edit2(NO1 As Long, bulan As String, tahun As String, npwp_kpp As String, _
                                kdPROYEK As String, kdCENTER As String, nama As String, npwp As String, _
                                NIK As String, alamat As String, Jabatan As String, P_L As String, _
                                ptkp As String, id1 As String, Insentif, Thr, Lainnya) As Integer
    
    '-- return
    '1: sukses edit
    '--
    
    Dim sql As String, t As String
    Dim return1 As Integer
    
    return1 = 0
    nama = cleanStr(nama)
    
    sql = "update pph21tahunan2 set No1 = '" & Trim(NO1) & "', Bulan = '" & Trim(bulan) & "', Tahun = '" & _
            Trim(tahun) & "', npwp_kpp = '" & Trim(npwp_kpp) & "', kdPROYEK = '" & _
            Trim(kdPROYEK) & "', kdCENTER = '" & Trim(kdCENTER) & "', nama = '" & Trim(nama) & "', npwp = '" & _
            Trim(npwp) & "', NIK = '" & Trim(NIK) & "', alamat = '" & Trim(alamat) & "', Jabatan = '" & _
            Trim(Jabatan) & "', P_L = '" & Trim(P_L) & "', Ptkp = '" & Trim(ptkp) & "', Insentif = '" & _
            Trim(Insentif) & "', THR = '" & Trim(Thr) & "', lainnya = '" & Trim(Lainnya) & "' where id1 = '" & _
            Trim(id1) & "'"
    
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbPph21Tahunan2_edit2 = -1
    Else
        tbPph21Tahunan2_edit2 = 1
    End If
End Function

Function tbPph_sap_delete(id1 As String) As Boolean
    Dim sql As String
    
    sql = "delete from pph_sap where id1 = '" & id1 & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        tbPph_sap_delete = False
    Else
        tbPph_sap_delete = True
    End If
End Function

Function tbAll2016_master_isKodeProyekLamaValid(kode_proyek_lama As String) As Boolean
    Dim sql As String
    
    If isDataAda("all2016_master", "kode_Proyek_lama", kode_proyek_lama, cnn) = True Then
        tbAll2016_master_isKodeProyekLamaValid = True
    Else
        tbAll2016_master_isKodeProyekLamaValid = False
    End If
End Function

Function tbAll2016_master_insert(NO1 As String, CABANG As String, DIVISI As String, _
                            NO_KONTRAK As String, NK_PPN As Currency, OWNER As String, _
                            PROYEK As String, KODE_ACPAC As String, _
                            kode_proyek_lama As String, kode_proyek_baru As String, _
                            DESCRIPTION As String, NPWP_OWNER As String) As String
    
    '-- return
    '1: sukses ""
    '--
    Dim klm(), isi()
    Dim nmTabel As String
    
    nmTabel = "all2016_master"
    klm = Array("NO", "CABANG", "divisi", "NO_KONTRAK", "NK_PPN", "OWNER", "PROYEK", _
                "KODE_ACPAC", "kode_Proyek_lama", "kode_Proyek_baru", "DESCRIPTION", "NPWP_OWNER")
    isi = Array(NO1, CABANG, DIVISI, NO_KONTRAK, NK_PPN, OWNER, PROYEK, KODE_ACPAC, _
                kode_proyek_lama, kode_proyek_baru, DESCRIPTION, NPWP_OWNER)
    
    If isDataAda2(cnn, nmTabel, "kode_Proyek_lama = '" & kode_proyek_lama & _
                "' and kode_Proyek_baru = '" & kode_proyek_baru & "'") = True Then
        
        'data sudah ada, update yuk
        If tbUpdate(nmTabel, klm, isi, cnn, "kode_Proyek_lama = '" & kode_proyek_lama & _
                "' and kode_Proyek_baru = '" & kode_proyek_baru & "'") = True Then
            tbAll2016_master_insert = "update"
        Else
            tbAll2016_master_insert = "error"
        End If
        
    Else
        
        If tbInsert(nmTabel, klm, isi, cnn) = True Then
            tbAll2016_master_insert = ""
        Else
            tbAll2016_master_insert = "error"
        End If
    End If
    
    
End Function


Function tbAll2016_maccount_insert(account As String, accpac As String, _
                            acct_name As String) As String
    
    '-- return
    '1: sukses ""
    '--
    Dim klm(), isi()
    Dim nmTabel As String
    
    nmTabel = "all2016_maccount"
    klm = Array("account", "accpac", "acct_name")
    isi = Array(account, accpac, acct_name)
    
    If isDataAda2(cnn, nmTabel, "account = '" & account & "' ") = True Then
        
        'data sudah ada, update yuk
        If tbUpdate(nmTabel, klm, isi, cnn, "account = '" & account & "' ") = True Then
            tbAll2016_maccount_insert = "update"
        Else
            tbAll2016_maccount_insert = "error"
        End If
        
    Else
        
        If tbInsert(nmTabel, klm, isi, cnn) = True Then
            tbAll2016_maccount_insert = ""
        Else
            tbAll2016_maccount_insert = "error"
        End If
    End If
    
    
End Function

Function tbAll2016_tb_insert(tahun As String, bulan As String, kode_proyek_lama As String, _
                            kode_proyek_baru As String, kode_akun As String, _
                            deskripsi As String, nama_proyek As String, debit As Currency, _
                            kredit As Currency, Icon As String, cek_data_ada As Boolean) As String
    
    '-- return
    '1: sukses ""
    '--
    Dim klm(), isi()
    Dim nmTabel As String
    
    nmTabel = "all2016_tb"
    klm = Array("tahun", "bulan", "kode_proyek_lama", "kode_proyek_baru", "kode_akun", _
                "deskripsi", "nama_proyek", "debit", "kredit", "icon")
    isi = Array(tahun, bulan, kode_proyek_lama, kode_proyek_baru, kode_akun, _
                deskripsi, nama_proyek, debit, kredit, Icon)
    
    If cek_data_ada = True Then
    
        If isDataAda2(cnn, nmTabel, "kode_Proyek_lama = '" & kode_proyek_lama & _
                    "' and kode_Proyek_baru = '" & kode_proyek_baru & "' and tahun = '" & tahun & _
                    "' and bulan = '" & bulan & "' and kode_akun = '" & kode_akun & "'") = True Then
            
            'data sudah ada, update yuk
            If tbUpdate(nmTabel, klm, isi, cnn, "kode_Proyek_lama = '" & kode_proyek_lama & _
                    "' and kode_Proyek_baru = '" & kode_proyek_baru & "' and tahun = '" & tahun & _
                    "' and bulan = '" & bulan & "' and kode_akun = '" & kode_akun & "'") = True Then
                tbAll2016_tb_insert = "update"
            Else
                tbAll2016_tb_insert = "error"
            End If
            
        Else
            
            If tbInsert(nmTabel, klm, isi, cnn) = True Then
                tbAll2016_tb_insert = ""
            Else
                tbAll2016_tb_insert = "error"
            End If
        End If
    Else
        'langsung insert
        If tbInsert(nmTabel, klm, isi, cnn) = True Then
            tbAll2016_tb_insert = ""
        Else
            tbAll2016_tb_insert = "error"
        End If
    End If
    
End Function

Function tbAll2016_fp_insert(kode_divisi As String, kode_proyek_lama As String, _
                        kode_proyek_baru As String, tahun As String, tgl_fp As Date, _
                        no_fp As String, dpp As Currency, ppn As Currency, _
                        KETERANGAN As String, cek_data_ada As Boolean) As String
    
    '-- return
    '1: sukses ""
    '--
    Dim klm(), isi()
    Dim nmTabel As String
    
    nmTabel = "all2016_fp"
    klm = Array("kode_divisi", "kode_proyek_lama", "kode_proyek_baru", "tahun", "tgl_fp", _
                "no_fp", "dpp", "ppn", "keterangan")
    isi = Array(kode_divisi, kode_proyek_lama, kode_proyek_baru, tahun, set_tgl_perv(tgl_fp), _
                no_fp, dpp, ppn, KETERANGAN)
    
    If cek_data_ada = True Then
    
        If isDataAda2(cnn, nmTabel, "kode_Proyek_lama = '" & kode_proyek_lama & _
                    "' and kode_Proyek_baru = '" & kode_proyek_baru & "' and tahun = '" & tahun & _
                    "' and no_fp = '" & no_fp & "' ") = True Then
            
            'data sudah ada, update yuk
            If tbUpdate(nmTabel, klm, isi, cnn, "kode_Proyek_lama = '" & kode_proyek_lama & _
                    "' and kode_Proyek_baru = '" & kode_proyek_baru & "' and tahun = '" & tahun & _
                    "' and no_fp = '" & no_fp & "' ") = True Then
                tbAll2016_fp_insert = "update"
            Else
                tbAll2016_fp_insert = "error"
            End If
            
        Else
            
            If tbInsert(nmTabel, klm, isi, cnn) = True Then
                tbAll2016_fp_insert = ""
            Else
                tbAll2016_fp_insert = "error"
            End If
        End If
    Else
        'langsung insert
        If tbInsert(nmTabel, klm, isi, cnn) = True Then
            tbAll2016_fp_insert = ""
        Else
            tbAll2016_fp_insert = "error"
        End If
    End If
    
End Function

Function tbAll2016_fp_insert2(kode_divisi As String, kode_proyek_lama As String, _
                        kode_proyek_baru As String, tahun As String, tgl_fp As Date, _
                        no_fp As String, dpp As Currency, ppn As Currency, _
                        KETERANGAN As String, cek_data_ada As Boolean, _
                        pk_pm As String, masa As String, npwp_rekanan As String, _
                        nama_rekanan As String, kode_fp As String) As String
    
    '-- ada tambahan 5 kolom baru
    '-- return
    '1: sukses ""
    '--
    Dim klm(), isi()
    Dim nmTabel As String
    
    nmTabel = "all2016_fp"
    klm = Array("kode_divisi", "kode_proyek_lama", "kode_proyek_baru", "tahun", "tgl_fp", _
                "no_fp", "dpp", "ppn", "keterangan", "pk_pm", "masa", _
                "npwp_rekanan", "nama_rekanan", "kode_fp")
    isi = Array(kode_divisi, kode_proyek_lama, kode_proyek_baru, tahun, set_tgl_perv(tgl_fp), _
                no_fp, dpp, ppn, KETERANGAN, pk_pm, masa, _
                npwp_rekanan, nama_rekanan, kode_fp)
    
    If cek_data_ada = True Then
    
        If isDataAda2(cnn, nmTabel, "kode_Proyek_lama = '" & kode_proyek_lama & _
                    "' and kode_Proyek_baru = '" & kode_proyek_baru & "' and tahun = '" & tahun & _
                    "' and no_fp = '" & no_fp & "' ") = True Then
            
            'data sudah ada, update yuk
            If tbUpdate(nmTabel, klm, isi, cnn, "kode_Proyek_lama = '" & kode_proyek_lama & _
                    "' and kode_Proyek_baru = '" & kode_proyek_baru & "' and tahun = '" & tahun & _
                    "' and no_fp = '" & no_fp & "' ") = True Then
                tbAll2016_fp_insert2 = "update"
            Else
                tbAll2016_fp_insert2 = "error"
            End If
            
        Else
            
            If tbInsert(nmTabel, klm, isi, cnn) = True Then
                tbAll2016_fp_insert2 = ""
            Else
                tbAll2016_fp_insert2 = "error"
            End If
        End If
    Else
        'langsung insert
        If tbInsert(nmTabel, klm, isi, cnn) = True Then
            tbAll2016_fp_insert2 = ""
        Else
            tbAll2016_fp_insert2 = "error"
        End If
    End If
    
End Function


Function tbebupot23_fp_insert(rsParam As ADODB.Recordset, _
                            kdDivisi As String, Optional cek_data_ada As Boolean = True) As String
    '-- return
    '1: sukses ""
    '--
    Dim klm(), isi()
    Dim nmTabel As String
    Dim a As Integer
    Dim npwp_kpp As String, Kode_Proyek As String, No_Bukti_Akuntansi As String
    Dim No_Faktur_Pajak  As String
    
    nmTabel = "ebupot23"
    
    klm = Array("NPWP_KPP", "Kode_Proyek", "No_Bukti_Akuntansi", "Jenis_Dokumen", "Tgl_Dokumen_ddMMyyyy", "No_Faktur_Pajak", _
"Kode_Form_Bukti_Potong", "Masa_Pajak", "Tahun_Pajak", "Pembetulan", "NPWP_WP_yang_Dipotong", "NIK_Yg_Dipotong", _
"Nomer_telepon", "Kode_Objek_Pajak", "Penanda_tangan_BP_Pengurus", "Mendapatkan_Fasilitas", "Nomor_SKB", _
"Nomor_Aturan_DTP", "NTPN_DTP", "Nama_WP_yang_Dipotong", "Alamat_WP_yang_Dipotong", "Nomor_Bukti_Potong", _
"Tanggal_Bukti_Potong", "Nilai_Bruto_1", "Tarif_1", "PPh_Yang_Dipotong__1", "Nilai_Bruto_2", "Tarif_2", _
"PPh_Yang_Dipotong__2", "Nilai_Bruto_3", "Tarif_3", "PPh_Yang_Dipotong__3", "Nilai_Bruto_4", "Tarif_4", _
"PPh_Yang_Dipotong__4", "Nilai_Bruto_5", "Tarif_5", "PPh_Yang_Dipotong__5", "Nilai_Bruto_6a_Nilai_Bruto_6", _
"Tarif_6a_Tarif_6", "PPh_Yang_Dipotong__6a_PPh_Yang_Dipotong__6", "Nilai_Bruto_6b_Nilai_Bruto_7", "Tarif_6b_Tarif_7", _
"PPh_Yang_Dipotong__6b_PPh_Yang_Dipotong__7", "Nilai_Bruto_6c_Nilai_Bruto_8", "Tarif_6c_Tarif_8", _
"PPh_Yang_Dipotong__6c_PPh_Yang_Dipotong__8", "Nilai_Bruto_9", "Tarif_9", "PPh_Yang_Dipotong__9", "Nilai_Bruto_10", _
"Perkiraan_Penghasilan_Netto10", "Tarif_10", "PPh_Yang_Dipotong__10", "Nilai_Bruto_11", _
"Perkiraan_Penghasilan_Netto11", "Tarif_11", "PPh_Yang_Dipotong__11", "Nilai_Bruto_12", _
"Perkiraan_Penghasilan_Netto12", "Tarif_12", "PPh_Yang_Dipotong__12", "Nilai_Bruto_13", "Tarif_13", _
"PPh_Yang_Dipotong__13", "Kode_Jasa_6d1_PMK_244_PMK03_2008", "Nilai_Bruto_6d1", "Tarif_6d1", "PPh_Yang_Dipotong__6d1", _
"Kode_Jasa_6d2_PMK_244_PMK03_2008", "Nilai_Bruto_6d2", "Tarif_6d2", "PPh_Yang_Dipotong__6d2", _
"Kode_Jasa_6d3_PMK_244_PMK03_2008", "Nilai_Bruto_6d3", "Tarif_6d3", "PPh_Yang_Dipotong__6d3", _
"Kode_Jasa_6d4_PMK_244_PMK03_2008", "Nilai_Bruto_6d4", "Tarif_6d4", "PPh_Yang_Dipotong__6d4", _
"Kode_Jasa_6d5_PMK_244_PMK03_2008", "Nilai_Bruto_6d5", "Tarif_6d5", "PPh_Yang_Dipotong__6d5", _
"Kode_Jasa_6d6_PMK_244_PMK03_2008", "Nilai_Bruto_6d6", "Tarif_6d6", "PPh_Yang_Dipotong__6d6", "Jumlah_Nilai_Bruto_", _
"Jumlah_PPh_Yang_Dipotong", "email", "kode_divisi")
               
    rsParam.MoveFirst
    ReDim isi(92)
    For a = 0 To rsParam.Fields.Count - 1
        If a = 2 Then
            isi(a) = cleanStr(rsParam.Fields(a))
        Else
            isi(a) = rsParam.Fields(a)
        End If
        
    Next
    isi(92) = kdDivisi
    
    
    'MsgBox (UBound(klm))
    If UBound(klm) <> UBound(isi) Then
        tbebupot23_fp_insert = "error array tidak sama"
        Exit Function
    End If
    
    npwp_kpp = cek_null(rsParam(0))
    Kode_Proyek = cek_null(rsParam(1))
    No_Bukti_Akuntansi = cleanStr(cek_null(rsParam(2)))
    No_Faktur_Pajak = cek_null(rsParam(5))
    
    If cek_data_ada = True Then
    
        If isDataAda2(cnn, nmTabel, "NPWP_KPP = '" & npwp_kpp & _
                    "' and Kode_Proyek = '" & Kode_Proyek & "' and No_Bukti_Akuntansi = '" & No_Bukti_Akuntansi & _
                    "' and No_Faktur_Pajak = '" & No_Faktur_Pajak & "' ") = True Then
            
            'data sudah ada, update yuk
            If tbUpdate(nmTabel, klm, isi, cnn, "NPWP_KPP = '" & npwp_kpp & _
                    "' and Kode_Proyek = '" & Kode_Proyek & "' and No_Bukti_Akuntansi = '" & No_Bukti_Akuntansi & _
                    "' and No_Faktur_Pajak = '" & No_Faktur_Pajak & "' ") = True Then
                tbebupot23_fp_insert = "update"
            Else
                tbebupot23_fp_insert = "error"
            End If
            
        Else
            
            If tbInsert(nmTabel, klm, isi, cnn) = True Then
                tbebupot23_fp_insert = ""
            Else
                tbebupot23_fp_insert = "error"
            End If
        End If
    Else
        'langsung insert
        If tbInsert(nmTabel, klm, isi, cnn) = True Then
            tbebupot23_fp_insert = ""
        Else
            tbebupot23_fp_insert = "error"
        End If
    End If
    
End Function

Function tbebupot26_insert(rsParam As ADODB.Recordset, _
                            kdDivisi As String, Optional cek_data_ada As Boolean = True) As String
    '-- return
    '1: sukses ""
    '--
    Dim klm(), isi()
    Dim rsNama As ADODB.Recordset
    Dim nmTabel As String
    Dim a As Integer, sql As String
    Dim npwp_kpp As String, Kode_Proyek As String, No_Bukti_Akuntansi As String
    Dim No_Faktur_Pajak  As String
    Dim TIN_  As String, Nama_WP_yang_Dipotong      As String, Jumlah_Nilai_Bruto_      As String
    
    nmTabel = "ebupot26"
    
    'get column name
    ReDim klm(96)
    sql = "select * from ebupot26 limit 1"
    If OpenRecordSet(cnn, rsNama, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        tbebupot26_insert = "error get kolom name"
        Exit Function
    End If
    For a = 1 To 96
        klm(a - 1) = rsNama.Fields(a).Name
    Next
    klm(96) = "kode_divisi"
                   
    'isi
    rsParam.MoveFirst
    ReDim isi(96)
    For a = 0 To rsParam.Fields.Count - 1
        If a = 2 Then
            isi(a) = cleanStr(rsParam.Fields(a))
        ElseIf a = 10 Then
            If cleanStr(rsParam.Fields(a)) = "0" Then
                isi(a) = "000000000000000"
            Else
                isi(a) = cleanStr(rsParam.Fields(a))
            End If
        Else
            isi(a) = rsParam.Fields(a)
        End If
        
    Next
    isi(96) = kdDivisi
    
    
    'MsgBox (UBound(klm))
    If UBound(klm) <> UBound(isi) Then
        tbebupot26_insert = "error array tidak sama"
        Exit Function
    End If
    
    npwp_kpp = cek_null(rsParam(0))
    Kode_Proyek = cek_null(rsParam(1))
    No_Bukti_Akuntansi = cleanStr(cek_null(rsParam(2)))
    No_Faktur_Pajak = cek_null(rsParam(5))
    TIN_ = cek_null(rsParam(11))
    Nama_WP_yang_Dipotong = cek_null(rsParam(14))
    Jumlah_Nilai_Bruto_ = cek_null(rsParam(93))
    
    If cek_data_ada = True Then
    
        If isDataAda2(cnn, nmTabel, "NPWP_KPP = '" & npwp_kpp & _
                    "' and Kode_Proyek = '" & Kode_Proyek & "' and No_Bukti_Akuntansi = '" & No_Bukti_Akuntansi & _
                    "' and No_Faktur_Pajak = '" & No_Faktur_Pajak & _
                    "' and TIN_ = '" & TIN_ & _
                    "' and Nama_WP_yang_Dipotong = '" & Nama_WP_yang_Dipotong & _
                    "' and Jumlah_Nilai_Bruto_ = '" & Jumlah_Nilai_Bruto_ & "' ") = True Then
            
            'data sudah ada, update yuk
            If tbUpdate(nmTabel, klm, isi, cnn, "NPWP_KPP = '" & npwp_kpp & _
                    "' and Kode_Proyek = '" & Kode_Proyek & "' and No_Bukti_Akuntansi = '" & No_Bukti_Akuntansi & _
                    "' and No_Faktur_Pajak = '" & No_Faktur_Pajak & _
                    "' and TIN_ = '" & TIN_ & _
                    "' and Nama_WP_yang_Dipotong = '" & Nama_WP_yang_Dipotong & _
                    "' and Jumlah_Nilai_Bruto_ = '" & Jumlah_Nilai_Bruto_ & "' ") = True Then
                tbebupot26_insert = "update"
            Else
                tbebupot26_insert = "error"
            End If
            
        Else
            
            If tbInsert(nmTabel, klm, isi, cnn) = True Then
                tbebupot26_insert = ""
            Else
                tbebupot26_insert = "error"
            End If
        End If
    Else
        'langsung insert
        If tbInsert(nmTabel, klm, isi, cnn) = True Then
            tbebupot26_insert = ""
        Else
            tbebupot26_insert = "error"
        End If
    End If
    
End Function

Sub tbebupot23_res_updateBP(JENIS_PPH As String, NOMOR_DOK_REF As String, _
                        Kode_Objek_Pajak As String, NOMOR_BUPOT As String, _
                        pph_dipotong As Currency)

    Dim sql As String
    
    'update di ebupot23/26
    If UCase(Trim(JENIS_PPH)) = "PPH23" Then
        sql = "update ebupot23 set Nomor_Bukti_Potong = '" & Trim(NOMOR_BUPOT) & _
            "', Jumlah_PPh_Yang_Dipotong = '" & pph_dipotong & _
            "' where No_Faktur_Pajak = '" & Trim(NOMOR_DOK_REF) & _
            "' and Kode_Objek_Pajak = '" & Trim(Kode_Objek_Pajak) & "'"
    ElseIf UCase(Trim(JENIS_PPH)) = "PPH26" Then
        sql = "update ebupot26 set Nomor_Bukti_Potong = '" & Trim(NOMOR_BUPOT) & _
            "', Jumlah_PPh_Yang_Dipotong = '" & pph_dipotong & _
            "' where No_Faktur_Pajak = '" & Trim(NOMOR_DOK_REF) & _
            "' and Kode_Objek_Pajak = '" & Trim(Kode_Objek_Pajak) & "'"
    End If
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("error", "", sql)
    End If
End Sub


Function tbebupot23_res_insert(rsParam As ADODB.Recordset, _
                            Optional cek_data_ada As Boolean = True) As String
    '-- return
    '1: sukses ""
    '--
    Dim klm(), isi()
    Dim rsNama As ADODB.Recordset
    Dim nmTabel As String
    Dim a As Integer, sql As String
    
    Dim JENIS_PPH As String, NOMOR_DOK_REF As String
    Dim Kode_Objek_Pajak As String, NOMOR_BUPOT As String, pph_dipotong As Currency
    
    nmTabel = "ebupot23_result"
    
    'get column name
    ReDim klm(15)
    sql = "select * from ebupot23_result limit 1"
    If OpenRecordSet(cnn, rsNama, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        tbebupot23_res_insert = "error get kolom name"
        Exit Function
    End If
    For a = 1 To 16
        klm(a - 1) = rsNama.Fields(a).Name
    Next
                   
    'isi
    rsParam.MoveFirst
    ReDim isi(15)
    For a = 0 To rsParam.Fields.Count - 1
        'If a = 2 Then
        '    isi(a) = cleanStr(rsParam.Fields(a))
        'ElseIf a = 10 Then
        '    If cleanStr(rsParam.Fields(a)) = "0" Then
        '        isi(a) = "000000000000000"
        '    Else
        '        isi(a) = cleanStr(rsParam.Fields(a))
        '    End If
        'Else
            isi(a) = rsParam.Fields(a)
        'End If
    Next
    
    'MsgBox (UBound(klm))
    If UBound(klm) <> UBound(isi) Then
        tbebupot23_res_insert = "error array tidak sama"
        Exit Function
    End If
    
    JENIS_PPH = cek_null(rsParam(0))
    NOMOR_BUPOT = cek_null(rsParam(3))
    Kode_Objek_Pajak = cek_null(rsParam(7))
    NOMOR_DOK_REF = cek_null(rsParam(13))
    pph_dipotong = cek_Money(rsParam(9))
    
                                
    
    If cek_data_ada = True Then
    
        If isDataAda2(cnn, nmTabel, "NOMOR_BUPOT = '" & NOMOR_BUPOT & _
                    "' and KODE_OBJEK_PAJAK = '" & Kode_Objek_Pajak & _
                    "' and NOMOR_DOK_REF = '" & NOMOR_DOK_REF & _
                    "' ") = True Then
            
            'data sudah ada, update yuk
            If tbUpdate(nmTabel, klm, isi, cnn, "NOMOR_BUPOT = '" & NOMOR_BUPOT & _
                    "' and KODE_OBJEK_PAJAK = '" & Kode_Objek_Pajak & _
                    "' and NOMOR_DOK_REF = '" & NOMOR_DOK_REF & _
                    "' ") = True Then
                tbebupot23_res_insert = "update"
            Else
                tbebupot23_res_insert = "error"
            End If
            
        Else
            
            If tbInsert(nmTabel, klm, isi, cnn) = True Then
                tbebupot23_res_insert = ""
            Else
                tbebupot23_res_insert = "error"
            End If
        End If
    Else
        'langsung insert
        If tbInsert(nmTabel, klm, isi, cnn) = True Then
            tbebupot23_res_insert = ""
        Else
            tbebupot23_res_insert = "error"
        End If
    End If
    
    Call tbebupot23_res_updateBP(JENIS_PPH, NOMOR_DOK_REF, Kode_Objek_Pajak, _
                                NOMOR_BUPOT, pph_dipotong)
    
End Function


Function tbAll2016_BP_insert(kode_divisi As String, kode_proyek_lama As String, _
                        kode_proyek_baru As String, tahun As String, tgl_BP As Date, _
                        no_BP As String, dpp As Currency, PPh As Currency, _
                        no_fp As String, KETERANGAN As String, cek_data_ada As Boolean) As String
    
    '-- return
    '1: sukses ""
    '--
    Dim klm(), isi()
    Dim nmTabel As String
    
    nmTabel = "all2016_bp"
    klm = Array("kode_divisi", "kode_proyek_lama", "kode_proyek_baru", "tahun", "tgl_bp", _
                "no_bp", "dpp", "pph", "no_fp", "keterangan")
    isi = Array(kode_divisi, kode_proyek_lama, kode_proyek_baru, tahun, set_tgl_perv(tgl_BP), _
                no_BP, dpp, PPh, no_fp, KETERANGAN)
    
    If cek_data_ada = True Then
    
        If isDataAda2(cnn, nmTabel, "kode_Proyek_lama = '" & kode_proyek_lama & _
                    "' and kode_Proyek_baru = '" & kode_proyek_baru & "' and tahun = '" & tahun & _
                    "' and no_bp = '" & no_BP & "' ") = True Then
            
            'data sudah ada, update yuk
            If tbUpdate(nmTabel, klm, isi, cnn, "kode_Proyek_lama = '" & kode_proyek_lama & _
                    "' and kode_Proyek_baru = '" & kode_proyek_baru & "' and tahun = '" & tahun & _
                    "' and no_bp = '" & no_BP & "' ") = True Then
                tbAll2016_BP_insert = "update"
            Else
                tbAll2016_BP_insert = "error"
            End If
            
        Else
            
            If tbInsert(nmTabel, klm, isi, cnn) = True Then
                tbAll2016_BP_insert = ""
            Else
                tbAll2016_BP_insert = "error"
            End If
        End If
    Else
        'langsung insert
        If tbInsert(nmTabel, klm, isi, cnn) = True Then
            tbAll2016_BP_insert = ""
        Else
            tbAll2016_BP_insert = "error"
        End If
    End If
    
End Function


Function tbTaxInqury_insert(NO1 As String, trans_reference_no As String, created_date As String, _
                        transaction_date As String, posting_date As String, billing_id As String, _
                        NTB As String, NTPN As String, STAN As String, tax_type As String, _
                        deposite_type As String, NPWP_number As String, tax_payer_name As String, _
                        City As String, WP_Address As String, NPWP_Payer As String, _
                        Payer_Name As String, Payer_Address As String, NOP As String, _
                        Tax_Period As String, SK_Number As String, _
                        Customer_Reference_No As String, Beneficiary_Email As String, _
                        Remark As String, Extended_Payment_Detail As String, Currency1 As String, _
                        Amount As Currency, Signature_ID As String, Signature_Name As String, _
                        Status As String, Reason As String, Optional cek_data_ada As Boolean = True) As Integer
    
    '-- return
    '1: sukses insert
    '2: update
    '0: error
    '--
    
    Dim klm(), isi()
    Dim nmTabel As String
    
    nmTabel = "tax_inquiry"
    klm = Array("no1", "trans_reference_no", "created_date", "transaction_date", "posting_date", _
                "billing_id", "NTB", "NTPN", "STAN", "tax_type", _
                "deposite_type", "NPWP_number", "tax_payer_name", "City", "WP_Address", _
                "NPWP_Payer", "Payer_Name", "Payer_Address", "NOP", "Tax_Period", _
                "SK_Number", "Customer_Reference_No", "Beneficiary_Email", "Remark", "Extended_Payment_Detail", _
                "Currency1", "Amount", "Signature_ID", "Signature_Name", "Status", _
                "Reason ")
    isi = Array(NO1, trans_reference_no, (created_date), (transaction_date), (posting_date), _
                billing_id, NTB, NTPN, STAN, tax_type, _
                deposite_type, NPWP_number, tax_payer_name, City, WP_Address, _
                NPWP_Payer, Payer_Name, Payer_Address, NOP, Tax_Period, _
                SK_Number, Customer_Reference_No, Beneficiary_Email, Remark, Extended_Payment_Detail, _
                Currency1, Amount, Signature_ID, Signature_Name, Status, _
                Reason)
    
    If cek_data_ada = True Then
    
        If isDataAda2(cnn, nmTabel, "trans_reference_no = '" & trans_reference_no & _
                        "' and billing_id = '" & billing_id & "' and NTPN = '" & NTPN & _
                        "' and Customer_Reference_No = '" & Customer_Reference_No & "'") = True Then
            
            'data sudah ada, update yuk
            If tbUpdate(nmTabel, klm, isi, cnn, "trans_reference_no = '" & trans_reference_no & _
                        "' and billing_id = '" & billing_id & "' and NTPN = '" & NTPN & _
                        "' and Customer_Reference_No = '" & Customer_Reference_No & "'") = True Then
                tbTaxInqury_insert = 2
            Else
                tbTaxInqury_insert = 0
            End If
            
        Else
            
            If tbInsert(nmTabel, klm, isi, cnn) = True Then
                tbTaxInqury_insert = 1
            Else
                tbTaxInqury_insert = 0
            End If
        End If
    Else
        'langsung insert
        If tbInsert(nmTabel, klm, isi, cnn) = True Then
            tbTaxInqury_insert = 1
        Else
            tbTaxInqury_insert = 0
        End If
    End If
    
End Function

Function tbPph_sap_insert(document_number As String, year_month1 As String, posting_key As String, _
                        document_header_text As String, account As String, acct_name As String, _
                        profit_center As String, document_date As String, posting_date As String, _
                        text As String, reference As String, amount_in_lc As Currency, _
                        vendor As String, name_cust_ven As String, reversed_with As String, _
                        cost_center As String, user_name As String, entry_date As Date) As String
    
    '-- return
    '1: sukses ""
    '--
    
    Dim sql As String, t As String
    Dim return1 As String
    
    return1 = 0
    sql = "select F_insert_pph_sap('" & document_number & "','" & year_month1 & "','" & posting_key & _
            "','" & document_header_text & "','" & account & "','" & acct_name & _
            "','" & profit_center & "','" & document_date & "','" & posting_date & _
            "','" & text & "','" & reference & "','" & amount_in_lc & _
            "','" & vendor & "','" & name_cust_ven & "','" & reversed_with & _
            "','" & cost_center & "','" & user_name & "','" & entry_date & "');"
    t = cari_data1(cnn, sql)
    tbPph_sap_insert = t
End Function

Sub tbAll2016_loadDivisi(ByRef cb As ComboBox, Optional wAll As Integer = 1)
    Dim sql As String
    
    sql = "select distinct cabang_divisi from all2016"
    Call Load_combo(cb, sql, cnn, False, , wAll)
End Sub

Sub tbAll2016_loadProyek(ByRef cb As ComboBox, Optional kdDivisi As String = "", _
                        Optional wAll As Integer = 1)
    Dim sql As String
    
    If Trim(kdDivisi) = "" Then
        sql = "select distinct kode_Proyek from all2016"
    Else
        sql = "select distinct kode_Proyek from all2016 where CABANG_DIVISI = '" & Trim(kdDivisi) & "'"
    End If
    
    
    Call Load_combo(cb, sql, cnn, False, , wAll)
End Sub

Sub tbPph_sap_load_Account(ByRef cb As ComboBox)
    Dim sql As String
    
    
    sql = "select concat(account, ' - ',acct_name) " & _
            "From pph_sap_maccount order by account"
    Call Load_combo(cb, sql, cnn)
End Sub

Sub tbPph_sap_load_year_month1(ByRef cb As ComboBox)
    Dim sql As String
    
    sql = "select distinct(year_month1) " & _
            "From pph_sap order by year_month1"
    Call Load_combo(cb, sql, cnn)
End Sub

Sub tbPph_sap_load_posting_key(ByRef cb As ComboBox)
    Dim sql As String
    
    sql = "select distinct(posting_key) " & _
            "From pph_sap order by posting_key"
    Call Load_combo(cb, sql, cnn)
End Sub

Function tbPph22_insert(npwp_kpp As String, k02 As String, Masa_Pajak As String, Tahun_Pajak As String, _
                        Pembetulan As String, npwp As String, Nama_NPWP As String, alamat As String, _
                        Nomor_Bukti_Potong As String, Tanggal_Bukti_Potong As Date, k35 As String, _
                        k36 As String, k37 As String, k38 As String, k39 As String, k40 As String, _
                        k41 As String, k42 As String, k43 As String, Nilai_DPP As Currency, Tarif As String, _
                        Nilai_PPh As Currency, k47 As String, k48 As String, k49 As String, k50 As String, _
                        j51 As String, j52 As String, kode_divisi As String, kd_proyek As String, nott As String, nofaktur As String) As Integer
    
    '-- return
    '1: sukses insert
    '2: update
    '--
    
    Dim sql As String
    Dim return1 As Integer
    
    return1 = 0
    If tbPphX_isDataAda(Tahun_Pajak, npwp_kpp, Nomor_Bukti_Potong, Pembetulan, "pph22") = True Then
        'data sudah ada, di hapus dulu
        If tbPphX_delete(Tahun_Pajak, npwp_kpp, Nomor_Bukti_Potong, Pembetulan, "pph22") = True Then
            return1 = 2
        Else
            tbPph22_insert = -1
            Exit Function
        End If
    End If
    
    Nama_NPWP = cleanStr(Nama_NPWP)
    sql = "insert into pph22(NPWP_KPP, k02, Masa_Pajak, " & _
            "Tahun_Pajak, Pembetulan, NPWP, " & _
            "Nama_NPWP, Alamat, Nomor_Bukti_Potong, " & _
            "Tanggal_Bukti_Potong, k35, k36, " & _
            "k37, k38, k39, " & _
            "k40, k41, k42, " & _
            "k43, Nilai_DPP, Tarif, " & _
            "Nilai_PPh, k47, k48, " & _
            "k49, k50, j51, " & _
            "j52, kode_divisi, tgl_import, " & _
            "kd_proyek, nott, nofaktur) values ('" & _
            Trim(npwp_kpp) & "','" & Trim(k02) & "','" & Trim(Masa_Pajak) & "','" & _
            Trim(Tahun_Pajak) & "','" & Trim(Pembetulan) & "','" & Trim(npwp) & "','" & _
            Trim(Nama_NPWP) & "','" & Trim(alamat) & "','" & Trim(Nomor_Bukti_Potong) & "','" & _
            set_tgl_perv(Tanggal_Bukti_Potong) & "','" & Trim(k35) & "','" & Trim(k36) & "','" & _
            Trim(k37) & "','" & Trim(k38) & "','" & Trim(k39) & "','" & _
            Trim(k40) & "','" & Trim(k41) & "','" & Trim(k42) & "','" & _
            Trim(k43) & "','" & Trim(Nilai_DPP) & "','" & Trim(Tarif) & "','" & _
            Trim(Nilai_PPh) & "','" & Trim(k47) & "','" & Trim(k48) & "','" & _
            Trim(k49) & "','" & Trim(k50) & "','" & Trim(j51) & "','" & _
            Trim(j52) & "','" & Trim(kode_divisi) & "','" & set_tgl_perv(Now) & "','" & _
            Trim(kd_proyek) & "','" & Trim(nott) & "','" & Trim(nofaktur) & "')"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbPph22_insert = -1
    Else
        If return1 = 2 Then
            tbPph22_insert = 2
        Else
            tbPph22_insert = 1
        End If
    End If
End Function

Function tbPph26_insert(npwp_kpp As String, Kode_Form As String, Masa_Pajak As String, Tahun_Pajak As String, _
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
                    kd_proyek As String, nott As String, nofaktur As String, email As String) As Integer
    
    
    
    '-- return
    '1: sukses insert
    '2: update
    '--
    
    Dim sql As String
    Dim return1 As Integer
    
    return1 = 0
    If tbPphX_isDataAda(Tahun_Pajak, npwp_kpp, Nomor_Bukti_Potong, Pembetulan, "pph26") = True Then
        'data sudah ada, di hapus dulu
        If tbPphX_delete(Tahun_Pajak, npwp_kpp, Nomor_Bukti_Potong, Pembetulan, "pph26") = True Then
            return1 = 2
        Else
            tbPph26_insert = -1
            Exit Function
        End If
    End If
    
    
    'yang di skip
    'Nilai_Bruto_9 s/d Nilai_Bruto_13
    'Kode_Jasa_6d2 s/d Kode_Jasa_6d6
    
    Nama_WP = cleanStr(Nama_WP)
    sql = "insert into pph26(NPWP_KPP, Kode_Form, Masa_Pajak, " & _
            "Tahun_Pajak, Pembetulan, NPWP_WP, " & _
            "Nama_WP, Alamat_WP, Nomor_Bukti_Potong, " & _
            "Tanggal_Bukti_Potong, " & _
            "Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1, " & _
            "Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, " & _
            "Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, " & _
            "Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, " & _
            "Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, " & _
            "Nilai_Bruto_6a, Tarif_6a, PPh_Yang_Dipotong__6a, " & _
            "Nilai_Bruto_6b, Tarif_6b, PPh_Yang_Dipotong__6b, " & _
            "Nilai_Bruto_6c, Tarif_6c, PPh_Yang_Dipotong__6c, " & _
            "Kode_Jasa_6d1, Nilai_Bruto_6d1, Tarif_6d1, PPh_Yang_Dipotong__6d1, "
    sql = sql & _
            "Jumlah_Nilai_Bruto_, Jumlah_PPh_Yang_Dipotong, kode_divisi, tgl_import, " & _
            "kd_proyek, nott, nofaktur, email) values ('" & _
            Trim(npwp_kpp) & "', '" & Trim(Kode_Form) & "','" & Trim(Masa_Pajak) & "','" & _
            Trim(Tahun_Pajak) & "','" & Trim(Pembetulan) & "','" & Trim(npwp_wp) & "','" & _
            Trim(Nama_WP) & "','" & Trim(Alamat_WP) & "','" & Trim(Nomor_Bukti_Potong) & "','" & _
            set_tgl_perv(Tanggal_Bukti_Potong) & "','" & _
            Trim(Nilai_Bruto_1) & "','" & Trim(Tarif_1) & "','" & Trim(PPh_Yang_Dipotong__1) & "','" & _
            Trim(Nilai_Bruto_2) & "','" & Trim(Tarif_2) & "','" & Trim(PPh_Yang_Dipotong__2) & "','" & _
            Trim(Nilai_Bruto_3) & "','" & Trim(Tarif_3) & "','" & Trim(PPh_Yang_Dipotong__3) & "','" & _
            Trim(Nilai_Bruto_4) & "','" & Trim(Tarif_4) & "','" & Trim(PPh_Yang_Dipotong__4) & "','" & _
            Trim(Nilai_Bruto_5) & "','" & Trim(Tarif_5) & "','" & Trim(PPh_Yang_Dipotong__5) & "','" & _
            Trim(Nilai_Bruto_6a) & "','" & Trim(Tarif_6a) & "','" & Trim(PPh_Yang_Dipotong__6a) & "','" & _
            Trim(Nilai_Bruto_6b) & "','" & Trim(Tarif_6b) & "','" & Trim(PPh_Yang_Dipotong__6b) & "','" & _
            Trim(Nilai_Bruto_6c) & "','" & Trim(Tarif_6c) & "','" & Trim(PPh_Yang_Dipotong__6c) & "','" & _
            Trim(Kode_Jasa_6d1) & "','" & Trim(Nilai_Bruto_6d1) & "','" & Trim(Tarif_6d1) & "','" & Trim(PPh_Yang_Dipotong__6d1) & "','" & _
            Trim(Jumlah_Nilai_Bruto_) & "','" & Trim(Jumlah_PPh_Yang_Dipotong) & "','" & Trim(kode_divisi) & "','" & set_tgl_perv(Now) & "','" & _
            Trim(kd_proyek) & "','" & Trim(nott) & "','" & Trim(nofaktur) & "', '" & Trim(email) & "')"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbPph26_insert = -1
    Else
        If return1 = 2 Then
            tbPph26_insert = 2
        Else
            tbPph26_insert = 1
        End If
    End If
End Function



Function tbPph42Konstruksi_insert(npwp_kpp As String, Kode_Form As String, Masa_Pajak As String, _
                                    Tahun_Pajak As String, Pembetulan As String, npwp_wp As String, _
                                    Nama_WP As String, Alamat_WP As String, Nomor_Bukti_Potong As String, _
                                    Tanggal_Bukti_Potong As Date, Jenis_Hadiah_Undian_1 As String, _
                                    Kode_Option_Tempat_Penyimpanan_1 As String, Jumlah_Nilai_Bruto_1 As Currency, _
                                    Tarif_1 As String, PPh_Yang_Dipotong__1 As Currency, _
                                    Jenis_Hadiah_Undian_2 As String, Kode_Option_Tempat_Penyimpanan_2 As String, _
                                    Jumlah_Nilai_Bruto_2 As Currency, Tarif_2 As String, _
                                    PPh_Yang_Dipotong__2 As Currency, Jenis_Hadiah_Undian_3 As String, _
                                    Kode_Option_Tempat_Penyimpanan_3 As String, Jumlah_Nilai_Bruto_3 As Currency, _
                                    Tarif_3 As String, PPh_Yang_Dipotong__3 As Currency, _
                                    Jenis_Hadiah_Undian_4 As String, Kode_Option_Tempat_Penyimpanan_4 As String, _
                                    Jumlah_Nilai_Bruto_4 As Currency, Tarif_4 As String, _
                                    PPh_Yang_Dipotong__4 As Currency, Jenis_Hadiah_Undian_5 As String, _
                                    Kode_Option_Tempat_Penyimpanan_5 As String, Jumlah_Nilai_Bruto_5 As Currency, _
                                    Tarif_5 As String, PPh_Yang_Dipotong__5 As Currency, _
                                    Jenis_Hadiah_Undian_6 As String, Jumlah_Nilai_Bruto_6 As Currency, _
                                    Tarif_6 As String, PPh_Yang_Dipotong__6 As Currency, _
                                    Jumlah_Nilai_Bruto_7 As Currency, Tarif_7 As String, _
                                    PPh_Yang_Dipotong_7 As Currency, Jenis_Penghasilan_8 As String, _
                                    Jumlah_Nilai_Bruto_8 As Currency, Tarif_8 As String, _
                                    PPh_Yang_Dipotong_8 As Currency, Jumlah_PPh_Yang_Dipotong As Currency, _
                                    Tanggal_Jatuh_Tempo_Obligasi As String, Tanggal_Perolehan_Obligasi As String, _
                                    Tanggal_Penjualan_Obligasi As String, Holding_Periode_Obligasi As String, _
                                    Time_Periode_Obligasi As String, kode_divisi As String, kd_proyek As String, nott As String, nofaktur As String, email As String) As Integer
    
    '-- return
    '1: sukses insert
    '2: update
    '--
    
    Dim sql As String
    Dim return1 As Integer
    
    return1 = 0
    If tbPphX_isDataAda(Tahun_Pajak, npwp_kpp, Nomor_Bukti_Potong, Pembetulan, "pph42_konstruksi") = True Then
        'data sudah ada, di hapus dulu
        If tbPphX_delete(Tahun_Pajak, npwp_kpp, Nomor_Bukti_Potong, Pembetulan, "pph42_konstruksi") = True Then
            return1 = 2
        Else
            tbPph42Konstruksi_insert = -1
            Exit Function
        End If
    End If
    
    
    'yang di skip
    
    sql = "insert into pph42_konstruksi(NPWP_KPP, Kode_Form, Masa_Pajak, " & _
            "Tahun_Pajak, Pembetulan, NPWP_WP, " & _
            "Nama_WP, Alamat_WP, Nomor_Bukti_Potong, " & _
            "Tanggal_Bukti_Potong, " & _
            "Jenis_Hadiah_Undian_1, Kode_Option_Tempat_Penyimpanan_1, " & _
            "Jumlah_Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1, " & _
            "Jenis_Hadiah_Undian_2, Kode_Option_Tempat_Penyimpanan_2, " & _
            "Jumlah_Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, " & _
            "Jenis_Hadiah_Undian_3, Kode_Option_Tempat_Penyimpanan_3, " & _
            "Jumlah_Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, " & _
            "Jenis_Hadiah_Undian_4, Kode_Option_Tempat_Penyimpanan_4," & _
            "Jumlah_Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, " & _
            "Jenis_Hadiah_Undian_5, Kode_Option_Tempat_Penyimpanan_5, " & _
            "Jumlah_Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, " & _
            "Jenis_Hadiah_Undian_6, Jumlah_Nilai_Bruto_6, Tarif_6, " & _
            "PPh_Yang_Dipotong__6, " & _
            "Jumlah_Nilai_Bruto_7, Tarif_7, PPh_Yang_Dipotong_7, " & _
            "Jenis_Penghasilan_8, Jumlah_Nilai_Bruto_8, Tarif_8, PPh_Yang_Dipotong_8, " & _
            "Jumlah_PPh_Yang_Dipotong, Tanggal_Jatuh_Tempo_Obligasi, Tanggal_Perolehan_Obligasi, " & _
            "Tanggal_Penjualan_Obligasi, Holding_Periode_Obligasi, Time_Periode_Obligasi, " & _
            "kode_divisi, tgl_import, kd_proyek, nott, nofaktur, email) values ('"
    sql = sql & _
            Trim(npwp_kpp) & "','" & Trim(Kode_Form) & "','" & Trim(Masa_Pajak) & "','" & _
            Trim(Tahun_Pajak) & "','" & Trim(Pembetulan) & "','" & Trim(npwp_wp) & "','" & _
            Trim(Nama_WP) & "','" & Trim(Alamat_WP) & "','" & Trim(Nomor_Bukti_Potong) & "','" & _
            set_tgl_perv(Tanggal_Bukti_Potong) & "','" & _
            Trim(Jenis_Hadiah_Undian_1) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_1) & "','" & _
            Trim(Jumlah_Nilai_Bruto_1) & "','" & Trim(Tarif_1) & "','" & Trim(PPh_Yang_Dipotong__1) & "','" & _
            Trim(Jenis_Hadiah_Undian_2) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_2) & "','" & _
            Trim(Jumlah_Nilai_Bruto_2) & "','" & Trim(Tarif_2) & "','" & Trim(PPh_Yang_Dipotong__2) & "','" & _
            Trim(Jenis_Hadiah_Undian_3) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_3) & "','" & _
            Trim(Jumlah_Nilai_Bruto_3) & "','" & Trim(Tarif_3) & "','" & Trim(PPh_Yang_Dipotong__3) & "','" & _
            Trim(Jenis_Hadiah_Undian_4) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_4) & "','" & _
            Trim(Jumlah_Nilai_Bruto_4) & "','" & Trim(Tarif_4) & "','" & Trim(PPh_Yang_Dipotong__4) & "','" & _
            Trim(Jenis_Hadiah_Undian_5) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_5) & "','" & _
            Trim(Jumlah_Nilai_Bruto_5) & "','" & Trim(Tarif_5) & "','" & Trim(PPh_Yang_Dipotong__5) & "','"
    sql = sql & _
            Trim(Jenis_Hadiah_Undian_6) & "','" & Trim(Jumlah_Nilai_Bruto_6) & "','" & Trim(Tarif_6) & "','" & _
            Trim(PPh_Yang_Dipotong__6) & "','" & _
            Trim(Jumlah_Nilai_Bruto_7) & "','" & Trim(Tarif_7) & "','" & Trim(PPh_Yang_Dipotong_7) & "','" & _
            Trim(Jenis_Penghasilan_8) & "','" & Trim(Jumlah_Nilai_Bruto_8) & "','" & Trim(Tarif_8) & "','" & Trim(PPh_Yang_Dipotong_8) & "','" & _
            Trim(Jumlah_PPh_Yang_Dipotong) & "','" & Trim(Tanggal_Jatuh_Tempo_Obligasi) & "','" & Trim(Tanggal_Perolehan_Obligasi) & "','" & _
            Trim(Tanggal_Penjualan_Obligasi) & "','" & Trim(Holding_Periode_Obligasi) & "','" & Trim(Time_Periode_Obligasi) & "','" & _
            Trim(kode_divisi) & "','" & set_tgl_perv(Now) & "','" & _
            Trim(kd_proyek) & "','" & Trim(nott) & "','" & Trim(nofaktur) & "', '" & Trim(email) & "')"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbPph42Konstruksi_insert = -1
    Else
        If return1 = 2 Then
            tbPph42Konstruksi_insert = 2
        Else
            tbPph42Konstruksi_insert = 1
        End If
    End If
End Function

Function tbPph42Sewa_insert(npwp_kpp As String, Kode_Form As String, Masa_Pajak As String, _
                                    Tahun_Pajak As String, Pembetulan As String, npwp_wp As String, _
                                    Nama_WP As String, Alamat_WP As String, Nomor_Bukti_Potong As String, _
                                    Tanggal_Bukti_Potong As Date, Jenis_Hadiah_Undian_1 As String, _
                                    Kode_Option_Tempat_Penyimpanan_1 As String, Jumlah_Nilai_Bruto_1 As Currency, _
                                    Tarif_1 As String, PPh_Yang_Dipotong__1 As Currency, _
                                    Jenis_Hadiah_Undian_2 As String, Kode_Option_Tempat_Penyimpanan_2 As String, _
                                    Jumlah_Nilai_Bruto_2 As Currency, Tarif_2 As String, _
                                    PPh_Yang_Dipotong__2 As Currency, Jenis_Hadiah_Undian_3 As String, _
                                    Kode_Option_Tempat_Penyimpanan_3 As String, Jumlah_Nilai_Bruto_3 As Currency, _
                                    Tarif_3 As String, PPh_Yang_Dipotong__3 As Currency, _
                                    Jenis_Hadiah_Undian_4 As String, Kode_Option_Tempat_Penyimpanan_4 As String, _
                                    Jumlah_Nilai_Bruto_4 As Currency, Tarif_4 As String, _
                                    PPh_Yang_Dipotong__4 As Currency, Jenis_Hadiah_Undian_5 As String, _
                                    Kode_Option_Tempat_Penyimpanan_5 As String, Jumlah_Nilai_Bruto_5 As Currency, _
                                    Tarif_5 As String, PPh_Yang_Dipotong__5 As Currency, _
                                    Jenis_Hadiah_Undian_6 As String, Jumlah_Nilai_Bruto_6 As Currency, _
                                    Tarif_6 As String, PPh_Yang_Dipotong__6 As Currency, _
                                    Jumlah_Nilai_Bruto_7 As Currency, Tarif_7 As String, _
                                    PPh_Yang_Dipotong_7 As Currency, Jenis_Penghasilan_8 As String, _
                                    Jumlah_Nilai_Bruto_8 As Currency, Tarif_8 As String, _
                                    PPh_Yang_Dipotong_8 As Currency, Jumlah_PPh_Yang_Dipotong As Currency, _
                                    Tanggal_Jatuh_Tempo_Obligasi As String, Tanggal_Perolehan_Obligasi As String, _
                                    Tanggal_Penjualan_Obligasi As String, Holding_Periode_Obligasi As String, _
                                    Time_Periode_Obligasi As String, kode_divisi As String, kd_proyek As String, nott As String, nofaktur As String, email As String) As Integer
    
    '-- return
    '1: sukses insert
    '2: update
    '--
    
    Dim sql As String
    Dim return1 As Integer
    
    return1 = 0
    If tbPphX_isDataAda(Tahun_Pajak, npwp_kpp, Nomor_Bukti_Potong, Pembetulan, "pph42_sewa") = True Then
        'data sudah ada, di hapus dulu
        If tbPphX_delete(Tahun_Pajak, npwp_kpp, Nomor_Bukti_Potong, Pembetulan, "pph42_sewa") = True Then
            return1 = 2
        Else
            tbPph42Sewa_insert = -1
            Exit Function
        End If
    End If
    
    
    'yang di skip
    Nama_WP = cleanStr(Nama_WP)
    sql = "insert into pph42_sewa(NPWP_KPP, Kode_Form, Masa_Pajak, " & _
            "Tahun_Pajak, Pembetulan, NPWP_WP, " & _
            "Nama_WP, Alamat_WP, Nomor_Bukti_Potong, " & _
            "Tanggal_Bukti_Potong, " & _
            "Jenis_Hadiah_Undian_1, Kode_Option_Tempat_Penyimpanan_1, " & _
            "Jumlah_Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1, " & _
            "Jenis_Hadiah_Undian_2, Kode_Option_Tempat_Penyimpanan_2, " & _
            "Jumlah_Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, " & _
            "Jenis_Hadiah_Undian_3, Kode_Option_Tempat_Penyimpanan_3, " & _
            "Jumlah_Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, " & _
            "Jenis_Hadiah_Undian_4, Kode_Option_Tempat_Penyimpanan_4," & _
            "Jumlah_Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, " & _
            "Jenis_Hadiah_Undian_5, Kode_Option_Tempat_Penyimpanan_5, " & _
            "Jumlah_Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, " & _
            "Jenis_Hadiah_Undian_6, Jumlah_Nilai_Bruto_6, Tarif_6, " & _
            "PPh_Yang_Dipotong__6, " & _
            "Jumlah_Nilai_Bruto_7, Tarif_7, PPh_Yang_Dipotong_7, " & _
            "Jenis_Penghasilan_8, Jumlah_Nilai_Bruto_8, Tarif_8, PPh_Yang_Dipotong_8, " & _
            "Jumlah_PPh_Yang_Dipotong, Tanggal_Jatuh_Tempo_Obligasi, Tanggal_Perolehan_Obligasi, " & _
            "Tanggal_Penjualan_Obligasi, Holding_Periode_Obligasi, Time_Periode_Obligasi, " & _
            "kode_divisi, tgl_import, kd_proyek, nott, nofaktur, email) values ('"
    sql = sql & _
            Trim(npwp_kpp) & "','" & Trim(Kode_Form) & "','" & Trim(Masa_Pajak) & "','" & _
            Trim(Tahun_Pajak) & "','" & Trim(Pembetulan) & "','" & Trim(npwp_wp) & "','" & _
            Trim(Nama_WP) & "','" & Trim(Alamat_WP) & "','" & Trim(Nomor_Bukti_Potong) & "','" & _
            set_tgl_perv(Tanggal_Bukti_Potong) & "','" & _
            Trim(Jenis_Hadiah_Undian_1) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_1) & "','" & _
            Trim(Jumlah_Nilai_Bruto_1) & "','" & Trim(Tarif_1) & "','" & Trim(PPh_Yang_Dipotong__1) & "','" & _
            Trim(Jenis_Hadiah_Undian_2) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_2) & "','" & _
            Trim(Jumlah_Nilai_Bruto_2) & "','" & Trim(Tarif_2) & "','" & Trim(PPh_Yang_Dipotong__2) & "','" & _
            Trim(Jenis_Hadiah_Undian_3) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_3) & "','" & _
            Trim(Jumlah_Nilai_Bruto_3) & "','" & Trim(Tarif_3) & "','" & Trim(PPh_Yang_Dipotong__3) & "','" & _
            Trim(Jenis_Hadiah_Undian_4) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_4) & "','" & _
            Trim(Jumlah_Nilai_Bruto_4) & "','" & Trim(Tarif_4) & "','" & Trim(PPh_Yang_Dipotong__4) & "','" & _
            Trim(Jenis_Hadiah_Undian_5) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_5) & "','" & _
            Trim(Jumlah_Nilai_Bruto_5) & "','" & Trim(Tarif_5) & "','" & Trim(PPh_Yang_Dipotong__5) & "','"
    sql = sql & _
            Trim(Jenis_Hadiah_Undian_6) & "','" & Trim(Jumlah_Nilai_Bruto_6) & "','" & Trim(Tarif_6) & "','" & _
            Trim(PPh_Yang_Dipotong__6) & "','" & _
            Trim(Jumlah_Nilai_Bruto_7) & "','" & Trim(Tarif_7) & "','" & Trim(PPh_Yang_Dipotong_7) & "','" & _
            Trim(Jenis_Penghasilan_8) & "','" & Trim(Jumlah_Nilai_Bruto_8) & "','" & Trim(Tarif_8) & "','" & Trim(PPh_Yang_Dipotong_8) & "','" & _
            Trim(Jumlah_PPh_Yang_Dipotong) & "','" & Trim(Tanggal_Jatuh_Tempo_Obligasi) & "','" & Trim(Tanggal_Perolehan_Obligasi) & "','" & _
            Trim(Tanggal_Penjualan_Obligasi) & "','" & Trim(Holding_Periode_Obligasi) & "','" & Trim(Time_Periode_Obligasi) & "','" & _
            Trim(kode_divisi) & "','" & set_tgl_perv(Now) & "','" & _
            Trim(kd_proyek) & "','" & Trim(nott) & "','" & Trim(nofaktur) & "', '" & Trim(email) & "')"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbPph42Sewa_insert = -1
    Else
        If return1 = 2 Then
            tbPph42Sewa_insert = 2
        Else
            tbPph42Sewa_insert = 1
        End If
    End If
End Function


Function tbPph42Obligasi_insert(npwp_kpp As String, Kode_Form As String, Masa_Pajak As String, _
                                    Tahun_Pajak As String, Pembetulan As String, npwp_wp As String, _
                                    Nama_WP As String, Alamat_WP As String, Nomor_Bukti_Potong As String, _
                                    Tanggal_Bukti_Potong As Date, Jenis_Hadiah_Undian_1 As String, _
                                    Kode_Option_Tempat_Penyimpanan_1 As String, Jumlah_Nilai_Bruto_1 As Currency, _
                                    Tarif_1 As String, PPh_Yang_Dipotong__1 As Currency, _
                                    Jenis_Hadiah_Undian_2 As String, Kode_Option_Tempat_Penyimpanan_2 As String, _
                                    Jumlah_Nilai_Bruto_2 As Currency, Tarif_2 As String, _
                                    PPh_Yang_Dipotong__2 As Currency, Jenis_Hadiah_Undian_3 As String, _
                                    Kode_Option_Tempat_Penyimpanan_3 As String, Jumlah_Nilai_Bruto_3 As Currency, _
                                    Tarif_3 As String, PPh_Yang_Dipotong__3 As Currency, _
                                    Jenis_Hadiah_Undian_4 As String, Kode_Option_Tempat_Penyimpanan_4 As String, _
                                    Jumlah_Nilai_Bruto_4 As Currency, Tarif_4 As String, _
                                    PPh_Yang_Dipotong__4 As Currency, Jenis_Hadiah_Undian_5 As String, _
                                    Kode_Option_Tempat_Penyimpanan_5 As String, Jumlah_Nilai_Bruto_5 As Currency, _
                                    Tarif_5 As String, PPh_Yang_Dipotong__5 As Currency, _
                                    Jenis_Hadiah_Undian_6 As String, Jumlah_Nilai_Bruto_6 As Currency, _
                                    Tarif_6 As String, PPh_Yang_Dipotong__6 As Currency, _
                                    Jumlah_Nilai_Bruto_7 As Currency, Tarif_7 As String, _
                                    PPh_Yang_Dipotong_7 As Currency, Jenis_Penghasilan_8 As String, _
                                    Jumlah_Nilai_Bruto_8 As Currency, Tarif_8 As String, _
                                    PPh_Yang_Dipotong_8 As Currency, Jumlah_PPh_Yang_Dipotong As Currency, _
                                    Tanggal_Jatuh_Tempo_Obligasi As String, Tanggal_Perolehan_Obligasi As String, _
                                    Tanggal_Penjualan_Obligasi As String, Holding_Periode_Obligasi As String, _
                                    Time_Periode_Obligasi As String, kode_divisi As String, kd_proyek As String, nott As String, nofaktur As String, email As String) As Integer
    
    '-- return
    '1: sukses insert
    '2: update
    '--
    
    Dim sql As String
    Dim return1 As Integer
    
    return1 = 0
    If tbPphX_isDataAda(Tahun_Pajak, npwp_kpp, Nomor_Bukti_Potong, Pembetulan, "pph42_obligasi") = True Then
        'data sudah ada, di hapus dulu
        If tbPphX_delete(Tahun_Pajak, npwp_kpp, Nomor_Bukti_Potong, Pembetulan, "pph42_obligasi") = True Then
            return1 = 2
        Else
            tbPph42Obligasi_insert = -1
            Exit Function
        End If
    End If
    
    
    'yang di skip
    Nama_WP = cleanStr(Nama_WP)
    sql = "insert into pph42_obligasi(NPWP_KPP, Kode_Form, Masa_Pajak, " & _
            "Tahun_Pajak, Pembetulan, NPWP_WP, " & _
            "Nama_WP, Alamat_WP, Nomor_Bukti_Potong, " & _
            "Tanggal_Bukti_Potong, " & _
            "Jenis_Hadiah_Undian_1, Kode_Option_Tempat_Penyimpanan_1, " & _
            "Jumlah_Nilai_Bruto_1, Tarif_1, PPh_Yang_Dipotong__1, " & _
            "Jenis_Hadiah_Undian_2, Kode_Option_Tempat_Penyimpanan_2, " & _
            "Jumlah_Nilai_Bruto_2, Tarif_2, PPh_Yang_Dipotong__2, " & _
            "Jenis_Hadiah_Undian_3, Kode_Option_Tempat_Penyimpanan_3, " & _
            "Jumlah_Nilai_Bruto_3, Tarif_3, PPh_Yang_Dipotong__3, " & _
            "Jenis_Hadiah_Undian_4, Kode_Option_Tempat_Penyimpanan_4," & _
            "Jumlah_Nilai_Bruto_4, Tarif_4, PPh_Yang_Dipotong__4, " & _
            "Jenis_Hadiah_Undian_5, Kode_Option_Tempat_Penyimpanan_5, " & _
            "Jumlah_Nilai_Bruto_5, Tarif_5, PPh_Yang_Dipotong__5, " & _
            "Jenis_Hadiah_Undian_6, Jumlah_Nilai_Bruto_6, Tarif_6, " & _
            "PPh_Yang_Dipotong__6, " & _
            "Jumlah_Nilai_Bruto_7, Tarif_7, PPh_Yang_Dipotong_7, " & _
            "Jenis_Penghasilan_8, Jumlah_Nilai_Bruto_8, Tarif_8, PPh_Yang_Dipotong_8, " & _
            "Jumlah_PPh_Yang_Dipotong, Tanggal_Jatuh_Tempo_Obligasi, Tanggal_Perolehan_Obligasi, " & _
            "Tanggal_Penjualan_Obligasi, Holding_Periode_Obligasi, Time_Periode_Obligasi, " & _
            "kode_divisi, tgl_import, kd_proyek, nott, nofaktur, email) values ('"
    sql = sql & _
            Trim(npwp_kpp) & "','" & Trim(Kode_Form) & "','" & Trim(Masa_Pajak) & "','" & _
            Trim(Tahun_Pajak) & "','" & Trim(Pembetulan) & "','" & Trim(npwp_wp) & "','" & _
            Trim(Nama_WP) & "','" & Trim(Alamat_WP) & "','" & Trim(Nomor_Bukti_Potong) & "','" & _
            set_tgl_perv(Tanggal_Bukti_Potong) & "','" & _
            Trim(Jenis_Hadiah_Undian_1) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_1) & "','" & _
            Trim(Jumlah_Nilai_Bruto_1) & "','" & Trim(Tarif_1) & "','" & Trim(PPh_Yang_Dipotong__1) & "','" & _
            Trim(Jenis_Hadiah_Undian_2) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_2) & "','" & _
            Trim(Jumlah_Nilai_Bruto_2) & "','" & Trim(Tarif_2) & "','" & Trim(PPh_Yang_Dipotong__2) & "','" & _
            Trim(Jenis_Hadiah_Undian_3) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_3) & "','" & _
            Trim(Jumlah_Nilai_Bruto_3) & "','" & Trim(Tarif_3) & "','" & Trim(PPh_Yang_Dipotong__3) & "','" & _
            Trim(Jenis_Hadiah_Undian_4) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_4) & "','" & _
            Trim(Jumlah_Nilai_Bruto_4) & "','" & Trim(Tarif_4) & "','" & Trim(PPh_Yang_Dipotong__4) & "','" & _
            Trim(Jenis_Hadiah_Undian_5) & "','" & Trim(Kode_Option_Tempat_Penyimpanan_5) & "','" & _
            Trim(Jumlah_Nilai_Bruto_5) & "','" & Trim(Tarif_5) & "','" & Trim(PPh_Yang_Dipotong__5) & "','"
    sql = sql & _
            Trim(Jenis_Hadiah_Undian_6) & "','" & Trim(Jumlah_Nilai_Bruto_6) & "','" & Trim(Tarif_6) & "','" & _
            Trim(PPh_Yang_Dipotong__6) & "','" & _
            Trim(Jumlah_Nilai_Bruto_7) & "','" & Trim(Tarif_7) & "','" & Trim(PPh_Yang_Dipotong_7) & "','" & _
            Trim(Jenis_Penghasilan_8) & "','" & Trim(Jumlah_Nilai_Bruto_8) & "','" & Trim(Tarif_8) & "','" & Trim(PPh_Yang_Dipotong_8) & "','" & _
            Trim(Jumlah_PPh_Yang_Dipotong) & "','" & Trim(Tanggal_Jatuh_Tempo_Obligasi) & "','" & Trim(Tanggal_Perolehan_Obligasi) & "','" & _
            Trim(Tanggal_Penjualan_Obligasi) & "','" & Trim(Holding_Periode_Obligasi) & "','" & Trim(Time_Periode_Obligasi) & "','" & _
            Trim(kode_divisi) & "','" & set_tgl_perv(Now) & "','" & _
            Trim(kd_proyek) & "','" & Trim(nott) & "','" & Trim(nofaktur) & "', '" & Trim(email) & "')"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbPph42Obligasi_insert = -1
    Else
        If return1 = 2 Then
            tbPph42Obligasi_insert = 2
        Else
            tbPph42Obligasi_insert = 1
        End If
    End If
End Function



Function tbSSPpph_insert(npwp_kpp As String, Kode_Form As String, Masa_Pajak_SSP As String, _
                        Tahun_Pajak_SSP As String, Pembetulan As String, NTPN As String, _
                        Tanggal_Setor_SSP As Date, Jumlah_SSP As Currency, Kode_KAP As String, _
                        Kode_Jenis_Setoran As String, Jenis_Pajak As String, kode_divisi As String) As Integer
    
    '-- return
    '1: sukses insert
    '2: update
    '--
    
    Dim sql As String
    Dim return1 As Integer
    
    return1 = 0
    If tbSSPpph_isDataAda(npwp_kpp, NTPN, Pembetulan) = True Then
        'data sudah ada, di hapus dulu
        If tbSSPpph_delete(npwp_kpp, NTPN, Pembetulan) = True Then
            return1 = 2
        Else
            tbSSPpph_insert = -1
            Exit Function
        End If
    End If
    
    
    'yang di skip
    
    sql = "insert into ssp_pph(NPWP_KPP, Kode_Form, Masa_Pajak, " & _
            "Tahun_Pajak, Pembetulan, NTPN, " & _
            "Tanggal_Setor_SSP, Jumlah_SSP, Kode_KAP, " & _
            "Kode_Jenis_Setoran, Jenis_Pajak, kode_divisi, tgl_import) values ('" & _
            Trim(npwp_kpp) & "','" & Trim(Kode_Form) & "','" & Trim(Masa_Pajak_SSP) & "','" & _
            Trim(Tahun_Pajak_SSP) & "','" & Trim(Pembetulan) & "','" & Trim(NTPN) & "','" & _
            set_tgl_perv(Tanggal_Setor_SSP) & "','" & Trim(Jumlah_SSP) & "','" & Trim(Kode_KAP) & "','" & _
            Trim(Kode_Jenis_Setoran) & "','" & Trim(Jenis_Pajak) & "','" & Trim(kode_divisi) & "','" & set_tgl_perv(Now) & "')"
    If ExecSQL1(cnn, sql) <> 0 Then
        sql = InputBox("", "", sql)
        tbSSPpph_insert = -1
    Else
        If return1 = 2 Then
            tbSSPpph_insert = 2
        Else
            tbSSPpph_insert = 1
        End If
    End If
End Function


Function tbSSPpph_isDataAda(npwp_kpp As String, NTPN As String, Pembetulan As String) As Boolean
    Dim sql As String, t As String
    
    sql = "select count(*) from ssp_pph where NPWP_KPP = '" & Trim(npwp_kpp) & _
            "' and NTPN = '" & Trim(NTPN) & _
            "' and Pembetulan = '" & Trim(Pembetulan) & "'"
    t = cari_data1(cnn, sql, True)
    If CInt(t) > 0 Then
        tbSSPpph_isDataAda = True
    Else
        tbSSPpph_isDataAda = False
    End If
End Function

Function tbSSPpph_delete(npwp_kpp As String, NTPN As String, Pembetulan As String) As Boolean
    Dim sql As String
    
    sql = "delete from ssp_pph where NPWP_KPP = '" & Trim(npwp_kpp) & _
            "' and NTPN = '" & Trim(NTPN) & _
            "' and Pembetulan = '" & Trim(Pembetulan) & "'"
    If ExecSQL1(cnn, sql) <> 0 Then
        tbSSPpph_delete = False
    Else
        tbSSPpph_delete = True
    End If
End Function

Sub fetch_dbRep_Divisi(kodeDivisi As String, ByRef sb1 As StatusBar)
    Dim cnnTemp As ADODB.connection
    Dim sql As String
    Dim rsSumber As ADODB.Recordset, rsTujuan As ADODB.Recordset
    Dim jRec As Long, c As Long, a As Integer
    
    'open cnnTemp
    Call db_Access_open(cnnTemp, App.Path & "\data\dbrep.mdb")
    
    'delete
    If Trim(kodeDivisi) = "ALL" Then
        sql = "delete from mdivisi"
    Else
        sql = "delete from mdivisi where kodedivisi = '" & Trim(kodeDivisi) & "'"
    End If
    
    If ExecSQL1(cnnTemp, sql) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    
    'insert
    If Trim(kodeDivisi) = "ALL" Then
        sql = "select * from mdivisi"
    Else
        sql = "select * from mdivisi where kodedivisi = '" & Trim(kodeDivisi) & "'"
    End If
    
    If OpenRecordSet(cnn, rsSumber, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    Else
        jRec = RecordCount(rsSumber)
        If jRec > 0 Then
            '-- open target
            sql = "select * from mdivisi"
            If OpenRecordSet(cnnTemp, rsTujuan, sql, adOpenDynamic, adLockPessimistic, adUseClient) <> 0 Then
                sql = InputBox("error", "", sql)
                Exit Sub
            End If
            '--
        
            rsSumber.MoveFirst
            c = 1
            Do While rsSumber.EOF = False
                Call info(2, "Fetch divisi. Run " & c & "/" & jRec, sb1)
                
                rsTujuan.AddNew
                For a = 0 To rsSumber.Fields.Count - 1
                    rsTujuan.Fields(a).Value = cek_null(rsSumber.Fields(a).Value)
                Next
                rsTujuan.Update
                
                rsSumber.MoveNext
                c = c + 1
            Loop
            Set rsSumber = Nothing
        End If
    End If
    
End Sub

Sub fetch_dbRep_tbAccpac(tahun As String, ByRef sb1 As StatusBar, ByRef stopLoad As Boolean)
    Dim cnnTemp As ADODB.connection
    Dim sql As String
    Dim rsSumber As ADODB.Recordset, rsTujuan As ADODB.Recordset
    Dim jRec As Long, c As Long, a As Integer
    Dim tahun2 As String
    
    'open cnnTemp
    Call db_Access_open(cnnTemp, App.Path & "\data\dbrep.mdb")
    
    'delete
    tahun2 = CStr(CLng(tahun) - 1)
    sql = "delete from tbaccpac where tahun = '" & Trim(tahun) & "' or tahun = '" & tahun2 & "'"
    
    If ExecSQL1(cnnTemp, sql) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    
    'insert
    If Trim(tahun) = "ALL" Then
        sql = "select * from tbaccpac"
    Else
        sql = "select * from tbaccpac where tahun = '" & Trim(tahun) & "' or tahun = '" & tahun2 & "'"
    End If
    
    If OpenRecordSet(cnn, rsSumber, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    Else
        jRec = RecordCount(rsSumber)
        If jRec > 0 Then
            '-- open target
            sql = "select * from tbaccpac"
            If OpenRecordSet(cnnTemp, rsTujuan, sql, adOpenDynamic, adLockPessimistic, adUseClient) <> 0 Then
                sql = InputBox("error", "", sql)
                Exit Sub
            End If
            '--
        
            rsSumber.MoveFirst
            c = 1
            Do While rsSumber.EOF = False
                Call info(2, "Fetch divisi. Run " & c & "/" & jRec, sb1)
                If stopLoad = True Then
                    Call pesan2("exit loop")
                    Exit Do
                End If
                
                rsTujuan.AddNew
                For a = 0 To rsSumber.Fields.Count - 1
                    rsTujuan.Fields(a).Value = cek_null(rsSumber.Fields(a).Value)
                Next
                rsTujuan.Update
                
                rsSumber.MoveNext
                c = c + 1
            Loop
            Set rsSumber = Nothing
        End If
    End If
    
End Sub

Sub fetch_dbRep_tbAccpac_subkon(tahun As String, ByRef sb1 As StatusBar, ByRef stopLoad As Boolean)
    Dim cnnTemp As ADODB.connection
    Dim sql As String
    Dim rsSumber As ADODB.Recordset, rsTujuan As ADODB.Recordset
    Dim jRec As Long, c As Long, a As Integer
    Dim kode As String, nmproyek As String, tahun2 As String
    Dim t01 As Currency, t02 As Currency, t03 As Currency, t04 As Currency
    Dim t05 As Currency, t06 As Currency, t07 As Currency, t08 As Currency
    Dim t09 As Currency, t10 As Currency, t11 As Currency, t12 As Currency
    Dim t13 As Currency, t14 As Currency, t15 As Currency, t16 As Currency
    Dim klm(), isi()
    
    'open cnnTemp
    Call db_Access_open(cnnTemp, App.Path & "\data\dbrep.mdb")
    
    'delete
    tahun2 = CStr(CLng(tahun) - 1)
    sql = "delete from tbaccpac_subkon where tahun = '" & Trim(tahun) & "'"
    
    If ExecSQL1(cnnTemp, sql) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    'list kode proyek dari data tahun sekarang
    sql = "select distinct kdproyek from tbaccpac where tahun = '" & tahun & "'"
    If OpenRecordSet(cnnTemp, rsSumber, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", , sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rsSumber)
    If jRec <= 0 Then Exit Sub
    rsSumber.MoveFirst
    c = 0
    Do While rsSumber.EOF = False
        c = c + 1
        Call info_progress(sb1, 1, c, jRec, "Load ekualisasi subkon")
        
        If stopLoad = True Then
            Call pesan2("exit load pph subkon")
            Exit Do
        End If
    
        kode = cek_null(rsSumber(0))
        nmproyek = tb_mProyek_getname(kode)
        
        't01
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) = '20111' " & _
                "and tahun = '" & tahun2 & "' and kdproyek = '" & kode & "'"
        t01 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t02
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) = '20112' " & _
                "and tahun = '" & tahun2 & "' and kdproyek = '" & kode & "'"
        t02 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t03
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) = '20113' " & _
                "and tahun = '" & tahun2 & "' and kdproyek = '" & kode & "'"
        t03 = CCur(cari_data1(cnnTemp, sql, True))
        
        '--- t04
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) = '20131' " & _
                "and tahun = '" & tahun2 & "' and kdproyek = '" & kode & "'"
        t04 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t05
        t05 = t01 + t02 + t03 + t04
        
        '-- t06
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('50301','50302') " & _
                "and tahun = '" & tahun & "' and kdproyek = '" & kode & "'"
        t06 = CCur(cari_data1(cnnTemp, sql, True))
        
        '--t07
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('50101') " & _
                "and tahun = '" & tahun & "' and kdproyek = '" & kode & "'"
        t07 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t08
        t08 = t06 + t07
        
        '-- t09
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('20111') " & _
                "and tahun = '" & tahun & "' and kdproyek = '" & kode & "'"
        t09 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t10
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('20112') " & _
                "and tahun = '" & tahun & "' and kdproyek = '" & kode & "'"
        t10 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t11
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('20113') " & _
                "and tahun = '" & tahun & "' and kdproyek = '" & kode & "'"
        t11 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t12
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('20131') " & _
                "and tahun = '" & tahun & "' and kdproyek = '" & kode & "'"
        t12 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t13
        t13 = t09 + t10 + t11 + t12
        
        '-- t14
        t14 = t05 + t08 - t13
        
        'insert
        klm = Array("kode", "nmproyek", "tahun", "t01", "t02", "t03", "t04", "t05", "t06", "t07", "t08", "t09", "t10", "t11", "t12", "t13", "t14")
        isi = Array(kode, nmproyek, tahun, t01, t02, t03, t04, t05, t06, t07, t08, t09, t10, t11, t12, t13, t14)
        If tbInsert("tbaccpac_subkon", klm, isi, cnnTemp) = True Then
        Else
            MsgBox "insert error", vbCritical
            Exit Sub
        End If
                
        rsSumber.MoveNext
    Loop
    
    
End Sub

Sub fetch_dbRep_tbAccpac_22(tahun As String, ByRef sb1 As StatusBar, ByRef stopLoad As Boolean)
    Dim cnnTemp As ADODB.connection
    Dim sql As String
    Dim rsSumber As ADODB.Recordset, rsTujuan As ADODB.Recordset
    Dim jRec As Long, c As Long, a As Integer
    Dim kode As String, nmproyek As String, tahun2 As String
    Dim t01 As Currency, t02 As Currency, t03 As Currency, t04 As Currency
    Dim t05 As Currency, t06 As Currency, t07 As Currency, t08 As Currency
    Dim t09 As Currency, t10 As Currency
    Dim klm(), isi()
    
    'open cnnTemp
    Call db_Access_open(cnnTemp, App.Path & "\data\dbrep.mdb")
    
    'delete
    tahun2 = CStr(CLng(tahun) - 1)
    sql = "delete from tbaccpac_22 where tahun = '" & Trim(tahun) & "'"
    
    If ExecSQL1(cnnTemp, sql) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    'list kode proyek dari data tahun sekarang
    sql = "select distinct kdproyek from tbaccpac where tahun = '" & tahun & "'"
    If OpenRecordSet(cnnTemp, rsSumber, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", , sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rsSumber)
    If jRec <= 0 Then Exit Sub
    rsSumber.MoveFirst
    c = 0
    Do While rsSumber.EOF = False
        c = c + 1
        Call info_progress(sb1, 1, c, jRec, "Load ekualisasi 22")
        
        If stopLoad = True Then
            Call pesan2("exit load pph21")
            Exit Do
        End If
    
        kode = cek_null(rsSumber(0))
        nmproyek = tb_mProyek_getname(kode)
        
        't01
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) = '20101' " & _
                "and tahun = '" & tahun2 & "' and kdproyek = '" & kode & "'"
        t01 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t02
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) = '20102' " & _
                "and tahun = '" & tahun2 & "' and kdproyek = '" & kode & "'"
        t02 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t03
        t03 = t01 + t02
        
        '--- t04
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('50201','50202') " & _
                "and tahun = '" & tahun & "' and kdproyek = '" & kode & "'"
        t04 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t05
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('20101') " & _
                "and tahun = '" & tahun & "' and kdproyek = '" & kode & "'"
        t05 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t06
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('20102') " & _
                "and tahun = '" & tahun & "' and kdproyek = '" & kode & "'"
        t06 = CCur(cari_data1(cnnTemp, sql, True))
        
        '--t07
        t07 = t05 + t06
        
        '-- t08
        t08 = t03 + t04 - t07
        
        '-- t09
        
        'insert
        klm = Array("kode", "nmproyek", "tahun", "t01", "t02", "t03", "t04", "t05", "t06", "t07", "t08")
        isi = Array(kode, nmproyek, tahun, t01, t02, t03, t04, t05, t06, t07, t08)
        If tbInsert("tbaccpac_22", klm, isi, cnnTemp) = True Then
        Else
            MsgBox "insert error", vbCritical
            Exit Sub
        End If
                
        rsSumber.MoveNext
    Loop
    
    
End Sub

Sub fetch_dbRep_tbAccpac_21(tahun As String, ByRef sb1 As StatusBar, ByRef stopLoad As Boolean)
    Dim cnnTemp As ADODB.connection
    Dim sql As String
    Dim rsSumber As ADODB.Recordset, rsTujuan As ADODB.Recordset
    Dim jRec As Long, c As Long, a As Integer
    Dim kode As String, nmproyek As String, tahun2 As String
    Dim t01 As Currency, t02 As Currency, t03 As Currency, t04 As Currency
    Dim t05 As Currency, t06 As Currency, t07 As Currency, t08 As Currency
    Dim t09 As Currency, t10 As Currency
    Dim klm(), isi()
    
    'open cnnTemp
    Call db_Access_open(cnnTemp, App.Path & "\data\dbrep.mdb")
    
    'delete
    tahun2 = CStr(CLng(tahun) - 1)
    sql = "delete from tbaccpac_21 where tahun = '" & Trim(tahun) & "'"
    
    If ExecSQL1(cnnTemp, sql) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    'list kode proyek dari data tahun sekarang
    sql = "select distinct kdproyek from tbaccpac where tahun = '" & tahun & "'"
    If OpenRecordSet(cnnTemp, rsSumber, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", , sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rsSumber)
    If jRec <= 0 Then Exit Sub
    rsSumber.MoveFirst
    c = 0
    Do While rsSumber.EOF = False
        c = c + 1
        Call info_progress(sb1, 1, c, jRec, "Load ekualisasi 21")
        
        If stopLoad = True Then
            Call pesan2("exit load pph21")
            Exit Do
        End If
    
        kode = cek_null(rsSumber(0))
        nmproyek = tb_mProyek_getname(kode)
        
        't01
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('51101','51201','80101','80201','80131') " & _
                "and tahun = '" & tahun & "' and kdproyek = '" & kode & "'"
        t01 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t02
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('51111','51112','51113','51114','51119','51213','51214'," & _
                "'51219','80111','80114','80133','80119','80213','80214','80219') " & _
                "and tahun = '" & tahun & "' and kdproyek = '" & kode & "'"
        t02 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t03
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('51115','51215','80115','80128','80136','80215') " & _
                "and tahun = '" & tahun & "' and kdproyek = '" & kode & "'"
        t03 = CCur(cari_data1(cnnTemp, sql, True))
        
        '--- t04
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('51121','51125','51221','51225','80121','80125','80221') " & _
                "and tahun = '" & tahun & "' and kdproyek = '" & kode & "'"
        t04 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t05
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('80124','80137') " & _
                "and tahun = '" & tahun & "' and kdproyek = '" & kode & "'"
        t05 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t06
        t06 = t01 + t02 + t03 + t04 + t05
        
        
        'insert
        klm = Array("kode", "nmproyek", "tahun", "t01", "t02", "t03", "t04", "t05", "t06", "t07", "t08")
        isi = Array(kode, nmproyek, tahun, t01, t02, t03, t04, t05, t06, t07, t08)
        If tbInsert("tbaccpac_21", klm, isi, cnnTemp) = True Then
        Else
            MsgBox "insert error", vbCritical
            Exit Sub
        End If
                
        rsSumber.MoveNext
    Loop
    
    
End Sub

Sub fetch_dbRep_tbAccpac_23(tahun As String, ByRef sb1 As StatusBar, ByRef stopLoad As Boolean)
    Dim cnnTemp As ADODB.connection
    Dim sql As String
    Dim rsSumber As ADODB.Recordset, rsTujuan As ADODB.Recordset
    Dim jRec As Long, c As Long, a As Integer
    Dim kode As String, nmproyek As String, tahun2 As String
    Dim t01 As Currency, t02 As Currency, t03 As Currency, t04 As Currency
    Dim t05 As Currency, t06 As Currency, t07 As Currency
    Dim klm(), isi()
    
    'open cnnTemp
    Call db_Access_open(cnnTemp, App.Path & "\data\dbrep.mdb")
    
    'delete
    tahun2 = CStr(CLng(tahun) - 1)
    sql = "delete from tbaccpac_23 where tahun = '" & Trim(tahun) & "'"
    
    If ExecSQL1(cnnTemp, sql) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    'list kode proyek dari data tahun sekarang
    sql = "select distinct kdproyek from tbaccpac where tahun = '" & tahun & "'"
    If OpenRecordSet(cnnTemp, rsSumber, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("sql error", , sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rsSumber)
    If jRec <= 0 Then Exit Sub
    rsSumber.MoveFirst
    c = 0
    Do While rsSumber.EOF = False
        c = c + 1
        Call info_progress(sb1, 1, c, jRec, "Load ekualisasi 23")
        
        If stopLoad = True Then
            Call pesan2("exit load pph23")
            Exit Do
        End If
    
        kode = cek_null(rsSumber(0))
        nmproyek = tb_mProyek_getname(kode)
        
        't01
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('50401','50402') " & _
                "and tahun = '" & tahun2 & "' and kdproyek = '" & kode & "'"
        t01 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t02
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('50431') " & _
                "and tahun = '" & tahun2 & "' and kdproyek = '" & kode & "'"
        t02 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t03
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('51801','51802','83211','83212') " & _
                "and tahun = '" & tahun2 & "' and kdproyek = '" & kode & "'"
        t03 = CCur(cari_data1(cnnTemp, sql, True))
        
        '--- t04
        sql = "select sum(debits + credits) " & _
                "From tbaccpac " & _
                "where left(accnum,5) in ('51852','51853','83252','83253') " & _
                "and tahun = '" & tahun2 & "' and kdproyek = '" & kode & "'"
        t04 = CCur(cari_data1(cnnTemp, sql, True))
        
        '-- t05
        t05 = t01 + t02 + t03 + t04
                
        'insert
        klm = Array("kode", "nmproyek", "tahun", "t01", "t02", "t03", "t04", "t05")
        isi = Array(kode, nmproyek, tahun, t01, t02, t03, t04, t05)
        If tbInsert("tbaccpac_23", klm, isi, cnnTemp) = True Then
        Else
            MsgBox "insert error", vbCritical
            Exit Sub
        End If
                
        rsSumber.MoveNext
    Loop
    
    
End Sub

Function tb_mProyek_getname(kd As String) As String
    Dim cnnTemp As ADODB.connection
    
    'open cnnTemp
    Call db_Access_open(cnnTemp, App.Path & "\data\dbrep.mdb")
    Dim sql As String, t As String
    
    sql = "select nm_proyek from mproyek where kd_proyek = '" & kd & "'"
    t = cari_data1(cnnTemp, sql)
    tb_mProyek_getname = (t)
End Function

Sub fetch_dbRep_KPP(npwp_kpp As String, ByRef sb1 As StatusBar)
    Dim cnnTemp As ADODB.connection
    Dim sql As String
    Dim rsSumber As ADODB.Recordset, rsTujuan As ADODB.Recordset
    Dim jRec As Long, c As Long, a As Integer
    
    'open cnnTemp
    Call db_Access_open(cnnTemp, App.Path & "\data\dbrep.mdb")
    
    'delete
    If Trim(npwp_kpp) = "ALL" Then
        sql = "delete from mkpp"
    Else
        sql = "delete from mkpp where npwp = '" & Trim(npwp_kpp) & "'"
    End If
    
    If ExecSQL1(cnnTemp, sql) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    
    'insert
    If Trim(npwp_kpp) = "ALL" Then
        sql = "select * from mkpp"
    Else
        sql = "select * from mkpp where npwp = '" & Trim(npwp_kpp) & "'"
    End If
    
    If OpenRecordSet(cnn, rsSumber, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    Else
        jRec = RecordCount(rsSumber)
        If jRec > 0 Then
            '-- open target
            sql = "select * from mkpp"
            If OpenRecordSet(cnnTemp, rsTujuan, sql, adOpenDynamic, adLockPessimistic, adUseClient) <> 0 Then
                sql = InputBox("error", "", sql)
                Exit Sub
            End If
            '--
        
            rsSumber.MoveFirst
            c = 1
            Do While rsSumber.EOF = False
                Call info(2, "Fetch divisi. Run " & c & "/" & jRec, sb1)
                
                rsTujuan.AddNew
                For a = 0 To rsSumber.Fields.Count - 1
                    rsTujuan.Fields(a).Value = cek_null(rsSumber.Fields(a).Value)
                Next
                rsTujuan.Update
                
                rsSumber.MoveNext
                c = c + 1
            Loop
            Set rsSumber = Nothing
        End If
    End If
    
End Sub

Sub fetch_dbRep_PPhX(npwp_kpp As String, kodeDivisi As String, kodeProyek As String, _
                    tahunPajak As String, masaPajak As String, _
                        ByRef sb1 As StatusBar, nmTabel1 As String, _
                        Optional jenisPPhSSP As String = "")
                        
    Dim cnnTemp As ADODB.connection
    Dim sql As String, kondisi As String
    Dim rsSumber As ADODB.Recordset, rsTujuan As ADODB.Recordset
    Dim jRec As Long, c As Long, a As Integer
    
    'open cnnTemp
    Call db_Access_open(cnnTemp, App.Path & "\data\dbrep.mdb")
    
    'delete
    sql = "delete from " & nmTabel1
    If ExecSQL1(cnnTemp, sql) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    
    'insert
    sql = "select * from " & nmTabel1
    
    kondisi = ""
    If Trim(npwp_kpp) = "ALL" Then
    Else
        kondisi = kondisi & " NPWP_KPP = '" & Trim(npwp_kpp) & "'"
    End If
    
    If Trim(kodeDivisi) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " kode_divisi = '" & Trim(kodeDivisi) & "'"
    End If
    
    If IsNumeric(kodeProyek) = True Then
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " kd_proyek = '" & Trim(kodeProyek) & "'"
    End If
    
    If Trim(tahunPajak) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Tahun_Pajak = '" & Trim(tahunPajak) & "'"
    End If
    
    If Trim(masaPajak) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Masa_Pajak = '" & Trim(masaPajak) & "'"
    End If
    
    If Trim(nmTabel1) = "ssp_pph" Then
        'hanya untuk ssp_pph
        If Trim(jenisPPhSSP) = "" Or Trim(jenisPPhSSP) = "ALL" Then
        Else
            If Trim(kondisi) = "" Then
            Else
                kondisi = kondisi & " AND "
            End If
            kondisi = kondisi & " Jenis_Pajak = '" & Trim(jenisPPhSSP) & "'"
        End If
    End If
    
    
    If Trim(kondisi) = "" Then
    Else
        sql = sql & " WHERE " & kondisi
    End If
    '---
    
    
    
    'MsgBox sql
    'sql = InputBox("error", "", sql)
    If OpenRecordSet(cnn, rsSumber, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    Else
        jRec = RecordCount2(rsSumber)
        If jRec > 0 Then
            '-- open target
            sql = "select * from " & nmTabel1
            If OpenRecordSet(cnnTemp, rsTujuan, sql, adOpenDynamic, adLockPessimistic, adUseClient) <> 0 Then
                sql = InputBox("error", "", sql)
                Exit Sub
            End If
            '--
        
            rsSumber.MoveFirst
            c = 1
            Do While rsSumber.EOF = False
                Call info_progress(sb1, 2, c, jRec, "dbRep " & nmTabel1)
                
                rsTujuan.AddNew
                ' index kolom di kurangi 2, karena di tabel transaksi, kolom paling akhir : email
                'For a = 0 To rsSumber.Fields.Count - 1
                For a = 0 To rsSumber.Fields.Count - 2
                    
                    'If a = 25 Then
                    '    MsgBox "a"
                    'End If
                    
                    If rsSumber.Fields(a).Type = adBigInt Or rsSumber.Fields(a).Type = adCurrency Or _
                        rsSumber.Fields(a).Type = adDecimal Or rsSumber.Fields(a).Type = adDouble Or _
                        rsSumber.Fields(a).Type = adInteger Or rsSumber.Fields(a).Type = adNumeric Then
                        
                        rsTujuan.Fields(a).Value = cek_null(rsSumber.Fields(a).Value, "0")
                    Else
                        rsTujuan.Fields(a).Value = cek_null(rsSumber.Fields(a).Value)
                    End If
                    
                Next
                rsTujuan.Update
                
                rsSumber.MoveNext
                c = c + 1
            Loop
            Set rsSumber = Nothing
        End If
    End If
    
End Sub


Sub fetch_dbRep_PPhX2(npwp_kpp As String, kodeDivisi As String, kodeProyek As String, _
                    tahunPajak As String, masaPajak As String, _
                        ByRef sb1 As StatusBar, nmTabel1 As String)
                        
    
    '-- fetch report untuk keperluan cetak bukti potong !!
    
    Dim cnnTemp As ADODB.connection
    Dim sql As String, kondisi As String, t As String
    Dim rsSumber As ADODB.Recordset, rsTujuan As ADODB.Recordset
    Dim jRec As Long, c As Long, a As Integer
    
    'open cnnTemp
    Call db_Access_open(cnnTemp, App.Path & "\data\dbrep.mdb")
    
    'delete
    sql = "delete from " & nmTabel1
    If ExecSQL1(cnnTemp, sql) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    End If
    
    
    'insert, select dulu sumber datanya
    sql = "select * from " & nmTabel1
    
    kondisi = ""
    If Trim(npwp_kpp) = "ALL" Then
    Else
        kondisi = kondisi & " NPWP_KPP = '" & Trim(npwp_kpp) & "'"
    End If
    
    If Trim(kodeDivisi) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " kode_divisi = '" & Trim(kodeDivisi) & "'"
    End If
    
    If IsNumeric(kodeProyek) = True Then
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " kd_proyek = '" & Trim(kodeProyek) & "'"
    End If
    
    If Trim(tahunPajak) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Tahun_Pajak = '" & Trim(tahunPajak) & "'"
    End If
    
    If Trim(masaPajak) = "ALL" Then
    Else
        If Trim(kondisi) = "" Then
        Else
            kondisi = kondisi & " AND "
        End If
        kondisi = kondisi & " Masa_Pajak = '" & Trim(masaPajak) & "'"
    End If
    
    If Trim(kondisi) = "" Then
    Else
        sql = sql & " WHERE " & kondisi
    End If
    '---
    
    
    
    'MsgBox sql
    'sql = InputBox("error", "", sql)
    If OpenRecordSet(cnn, rsSumber, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error", "", sql)
        Exit Sub
    Else
        jRec = RecordCount2(rsSumber)
        If jRec > 0 Then
            '-- open target
            sql = "select * from " & nmTabel1
            If OpenRecordSet(cnnTemp, rsTujuan, sql, adOpenDynamic, adLockPessimistic, adUseClient) <> 0 Then
                sql = InputBox("error", "", sql)
                Exit Sub
            End If
            '--
        
            rsSumber.MoveFirst
            c = 1
            Do While rsSumber.EOF = False
                Call info_progress(sb1, 2, c, jRec, "dbRepX2 " & nmTabel1)
                
                If c Mod 50 = 0 Then
                    If dbMySQL_open = False Then
                        Exit Sub
                    End If
                End If
                
                rsTujuan.AddNew
                For a = 0 To rsSumber.Fields.Count - 2
                    
                    If rsSumber.Fields(a).Type = adBigInt Or rsSumber.Fields(a).Type = adCurrency Or _
                        rsSumber.Fields(a).Type = adDecimal Or rsSumber.Fields(a).Type = adDouble Or _
                        rsSumber.Fields(a).Type = adInteger Or rsSumber.Fields(a).Type = adNumeric Then
                        
                        rsTujuan.Fields(a).Value = cek_null(rsSumber.Fields(a).Value, "0")
                    Else
                        rsTujuan.Fields(a).Value = cek_null(rsSumber.Fields(a).Value)
                    End If
                    
                    'terbilang, sesuai table
                    If nmTabel1 = "pph15" Then
                        rsTujuan(23) = UCase(angka2word(cek_Money(rsSumber.Fields(17).Value))) & " RUPIAH"
                    ElseIf nmTabel1 = "pph22" Then
                        rsTujuan(60) = UCase(angka2word(cek_Money(rsSumber.Fields(48).Value))) & " RUPIAH"
                    ElseIf nmTabel1 = "pph23" Then
                        rsTujuan(86) = UCase(angka2word(cek_Money(rsSumber.Fields(80).Value))) & " RUPIAH"
                    ElseIf nmTabel1 = "pph42_konstruksi" Then
                        rsTujuan(60) = UCase(angka2word(cek_Money(rsSumber.Fields(49).Value))) & " RUPIAH"
                    ElseIf nmTabel1 = "pph42_sewa" Then
                        rsTujuan(60) = UCase(angka2word(cek_Money(rsSumber.Fields(49).Value))) & " RUPIAH"
                    End If
                Next
                
                rsTujuan("namakpp").Value = tbMKpp_getNamaKPP(cek_null(rsSumber(0).Value))
                rsTujuan("kotakpp").Value = tbMKpp_getKotaKPP(cek_null(rsSumber(0).Value))
                
                rsTujuan.Update
                
                rsSumber.MoveNext
                c = c + 1
            Loop
            Set rsSumber = Nothing
        End If
    End If
    
End Sub

Sub fetch_Bukti_Potong(tahun As String, ByRef rs As ADODB.Recordset, ByRef cnnTemp As ADODB.connection, _
                        ByRef sb1 As StatusBar, ByRef adaError As Boolean)
    Dim id1 As String
    Dim rsKPP As ADODB.Recordset
    Dim sql As String
    Dim jRec As Long, c As Long
    
    Dim t As String, t2 As String
    Dim penghasilanKenaPajakSetahun As Currency
    Dim npwp_kpp As String
    
    Dim alamat As String, bulan_akhir As String, bulan_awal As String, Jabatan As String
    Dim jenis_kelamin As String, nama As String, Nama_Pemotong As String, nama_ttd As String
    Dim NIK As String, no_1 As Currency, no_10 As Currency, no_11 As Currency, no_12 As Currency
    Dim no_13 As Currency, no_14 As Currency, no_15 As Currency, no_16 As Currency, no_17 As Currency
    Dim no_18 As Currency, no_19 As Currency, no_2 As Currency, no_20 As Currency, no_3 As Currency
    Dim no_4 As Currency, no_5 As Currency, no_6 As Currency, no_7 As Currency, no_8 As Currency
    Dim no_9 As Currency, nomor As String, npwp As String
    Dim npwp_ttd As String, ptkp As String
    
    Dim no_13s As Currency, no_19s As Currency, jmlBulan As Integer
    Dim kdCENTER As String
    Dim penghasilan_netto_sblmnya As Currency, pph21_terutang_sblmnya As Currency
    
    On Error GoTo er1
    
    adaError = False
    
    
    '-----
    'pph21tahunan2
        '0  No1, Bulan, Tahun, " & _
        '3  NPWP_KPP, kdPROYEK, kdCENTER, " & _
        '6  Nama, NPWP, NIK, " & _
        '9  Alamat, Jabatan, P_L, " & _
        '12 PTKP, Gaji, Tnj_PPh, " & _
        '15 Tunjangan_Lain, JHT_JPN, Bruto, " & _
        '18 Insentif, THR, Lainnya, " & _
        '21 Pensiun_Potongan_Lain, id1, tglupdate
        '24 tunjangan_jab
    '----------
    npwp = cek_null(rs(7))
    nama = cek_null(rs(6))
    NIK = cek_null(rs(8))
    
    'id1 = cek_null(rs(22))
    id1 = tbPph21Tahunan2_get_ID1(npwp, NIK, nama, tahun)
    
    '-- cek di temporary, jika data sudah ada, exit
    sql = "select count(*) from buktipotong where npwp = '" & CekPetik(npwp) & _
            "' and nama = '" & CekPetik(nama) & "' and NIK = '" & CekPetik(NIK) & "' " & _
            " and tahun = '" & tahun & "'"
    t = cari_data1(cnnTemp, sql, True)
    If CInt(t) > 0 Then Exit Sub
    '----
    
    sql = "select NPWP_KPP from v_dis_npwpkpp where NPWP = '" & Trim(npwp) & _
            "' and NIK = '" & Trim(NIK) & "' and Nama = '" & Trim(nama) & "' and Tahun = '" & _
            Trim(tahun) & "' "
    
    If OpenRecordSet(cnn, rsKPP, sql, adOpenStatic, adLockReadOnly, adUseClient) <> 0 Then
        sql = InputBox("error sql", "", sql)
        Exit Sub
    End If
    
    jRec = RecordCount(rsKPP)
    If jRec <= 0 Then
        MsgBox "tidak ada data KPP", vbInformation
        Exit Sub
    Else
        'Call pesan2("data di " & jRec & " KPP", 1)
    End If
    
    If jRec > 1 Then
        c = 1
    End If
    
    rsKPP.MoveFirst
    c = 1
    no_13s = 0
    no_19s = 0
    Do While rsKPP.EOF = False
        Call info(2, "Load data. Run " & c & "/" & jRec, sb1)
        npwp_kpp = cek_null(rsKPP(0))
            
        '---
        
        Call tbPph21Tahunan2_getData_byId2(id1, nomor, alamat, jenis_kelamin, ptkp)
        
        nomor = tbPph21Tahunan2_getNomorBukti(npwp, NIK, nama, tahun, npwp_kpp)
        bulan_awal = adddigit(CLng(tbPph21Tahunan2_getBulanAwal(npwp, NIK, nama, tahun, npwp_kpp)), 2)
        bulan_akhir = adddigit(tbPph21Tahunan2_getBulanAkhir(npwp, NIK, nama, tahun, npwp_kpp), 2)
        'update bulan akhir
        nomor = Left(nomor, 4) & bulan_akhir & "." & Right(tahun, 2) & "." & Right(nomor, 7)
        jmlBulan = (CInt(bulan_akhir) - CInt(bulan_awal)) + 1
    
        'npwp_kpp = ok
        Nama_Pemotong = UCase(tbMKpp_get("nama", npwp_kpp))
        'npwp = OK
        'NIK = ok
        nama = UCase(nama)
        alamat = Trim(Left(alamat, 50))
        Jabatan = tbPph21Tahunan2_getJabatan(npwp, NIK, nama, tahun)
    
        'totalbruto = sum(Gaji + Tnj_PPh + JHT_JPN + Tunjangan_Lain + Insentif)
        Call tbPph21Tahunan2_getTotal2(npwp, NIK, nama, tahun, npwp_kpp, no_1, no_2, no_3, no_5, no_7 _
                                        , no_10, penghasilan_netto_sblmnya, pph21_terutang_sblmnya)
        
        'no_1 = cek_Money(tbPph21Tahunan2_getTotal("Gaji", npwp, NIK, nama, Tahun, npwp_kpp))
        'no_2 = cek_Money(tbPph21Tahunan2_getTotal("Tnj_PPh", npwp, NIK, nama, Tahun, npwp_kpp))
        'no_3 = cek_Money(tbPph21Tahunan2_getTotal("Tunjangan_Lain", npwp, NIK, nama, Tahun, npwp_kpp))
        no_4 = 0
        'no_5 = cek_Money(tbPph21Tahunan2_getTotal("JHT_JPN", npwp, NIK, nama, Tahun, npwp_kpp))
        no_6 = 0
        'no_7 = cek_Money(tbPph21Tahunan2_getTotal("Insentif + THR + Lainnya", npwp, NIK, nama, Tahun, npwp_kpp))
        no_8 = no_1 + no_2 + no_3 + no_4 + no_5 + no_6 + no_7
            
        no_9 = cek_Money(tbPph21Tahunan2_getTotalBiayaJabatan(npwp, NIK, nama, tahun, npwp_kpp))
        'no_10 = cek_Money(tbPph21Tahunan2_getTotal("Pensiun_Potongan_Lain", npwp, NIK, nama, Tahun, npwp_kpp))
        no_11 = no_9 + no_10
        no_12 = no_8 - no_11
        no_13 = no_13s
        
        '--- cek jika ada data penghasilan dari perusahaan sebelumnya
        If c = jRec Or c = 1 Then
            If penghasilan_netto_sblmnya > 0 Then
                no_13 = no_13 + penghasilan_netto_sblmnya
            End If
        End If
        '-------------
        
        'jRec = jumlah KPP
        If jRec > 1 Then
            If c = 1 Then
                no_14 = Round(no_12 * (12 / jmlBulan), 0)
            Else
                no_14 = no_12 + no_13
            End If
        Else
            no_14 = no_12 + no_13
        End If
        
        no_13s = no_13s + no_12
        
        no_15 = tbM_Ptkp_getNilai(ptkp)
            
        If no_14 - no_15 > 0 Then
            no_16 = NearestThousand(no_14 - no_15)
        Else
            no_16 = 0
        End If
    
        penghasilanKenaPajakSetahun = no_16
    
        no_17 = cek_Money(get_pph21Setahun(penghasilanKenaPajakSetahun))
        If Left(cleanNpwp(npwp), 8) = "00000000" Then
            no_17 = no_17 * 1.2
        End If
        
        no_18 = no_19s
        
        '--- cek jika ada data pph dari perusahaan sebelumnya
        If c = jRec Then
            If pph21_terutang_sblmnya > 0 Then
                no_18 = no_18 + pph21_terutang_sblmnya
            End If
        End If
        '-------------
        
        If jRec > 1 Then
            If c = 1 Then
                no_19 = Round((no_17 / 12) * jmlBulan, 0)
            Else
                no_19 = no_17 - no_18
            End If
        Else
            no_19 = no_17 - no_18
        End If
        If no_19 <= 0 Then no_19 = 0
        
        'no_19s = no_19
        no_19s = no_19s + no_19
        
        no_20 = no_19
        
        npwp_ttd = "09.321.683.6411000"
        nama_ttd = "FARID FACHRUR RAZI"
        kdCENTER = tbPph21Tahunan2_get_kdCENTER(npwp, NIK, nama, tahun, npwp_kpp, bulan_akhir)
        
        '--insert ke dbtemp
        '- jika di panggil melalui form rekap, insert ke db utama
        sql = "insert into buktipotong (nomor, tahun, bulan_awal, " & _
                "bulan_akhir, npwp_pemotong, nama_pemotong, " & _
                "npwp, NIK, Nama, " & _
                "Alamat, Jenis_kelamin, ptkp, " & _
                "jabatan, no_1, no_2, " & _
                "no_3, no_4, no_5, " & _
                "no_6, no_7, no_8, " & _
                "no_9, no_10, no_11, " & _
                "no_12, no_13, no_14, " & _
                "no_15, no_16, no_17, " & _
                "no_18, no_19, no_20, " & _
                "npwp_ttd, nama_ttd, kdCENTER) values ('" & _
                Trim(nomor) & "','" & Trim(tahun) & "','" & Trim(bulan_awal) & "','" & _
                Trim(bulan_akhir) & "','" & Trim(npwp_kpp) & "','" & Trim(Nama_Pemotong) & "','" & _
                Trim(npwp) & "','" & Trim(NIK) & "','" & Trim(nama) & "','" & _
                Trim(alamat) & "','" & Trim(jenis_kelamin) & "','" & Trim(ptkp) & "','" & _
                Trim(Jabatan) & "','" & Trim(no_1) & "','" & Trim(no_2) & "','" & _
                Trim(no_3) & "','" & Trim(no_4) & "','" & Trim(no_5) & "','" & _
                Trim(no_6) & "','" & Trim(no_7) & "','" & Trim(no_8) & "','" & _
                Trim(no_9) & "','" & Trim(no_10) & "','" & Trim(no_11) & "','" & _
                Trim(no_12) & "','" & Trim(no_13) & "','" & Trim(no_14) & "','" & _
                Trim(no_15) & "','" & Trim(no_16) & "','" & Trim(no_17) & "','" & _
                Trim(no_18) & "','" & Trim(no_19) & "','" & Trim(no_20) & "','" & _
                Trim(npwp_ttd) & "','" & Trim(nama_ttd) & "', '" & kdCENTER & "')"
        If ExecSQL1(cnnTemp, sql) <> 0 Then
            sql = InputBox("sql error", "", sql)
            adaError = True
            Exit Do
        End If
        '---
        rsKPP.MoveNext
        c = c + 1
    Loop
    Set rsKPP = Nothing
    
    Exit Sub
er1:
    MsgBox Err.DESCRIPTION, vbCritical
End Sub

Function tbVariabel_get(key1 As String) As String
    Dim t As String
    
    If dbMySQL_open = True Then
        t = tbGet(cnn, "tvariabel", "ket", "key1 = '" & key1 & "'")
    Else
        Call pesan2("open DBOnline failed")
    End If
    tbVariabel_get = t
End Function

Sub tbVariabel_set(key1 As String, value1 As String)
    Dim t As String
    Dim kolom(), isi()
    
    kolom = Array("key1", "ket")
    isi = Array(key1, value1)
    
    If dbMySQL_open = True Then
        If isDataAda("tvariabel", "key1", key1, cnn) = True Then
            Call tbUpdate("tvariabel", kolom, isi, cnn, "key1 = '" & key1 & "'")
        Else
            Call tbInsert("tvariabel", kolom, isi, cnn)
        End If
    Else
        Call pesan2("open DBOnline failed")
    End If
End Sub

'--- ubni_direct
Function tbBniDirect_insert(kodeDivisi As String, noRek As String, _
                        nmPemegang As String, user_bni) As Boolean

    Dim klm(), isi()
    Dim nmTabel As String
    
    nmTabel = "ubni_direct"
    
    klm = Array("kodedivisi", "norek", "nmpemegang", "user_bni")
    isi = Array(kodeDivisi, noRek, nmPemegang, user_bni)
    If tbInsert(nmTabel, klm, isi, cnn) = True Then
        tbBniDirect_insert = True
    Else
        tbBniDirect_insert = False
    End If
End Function

Function tbBniDirect_update(kodeDivisi As String, noRek As String, _
                        nmPemegang As String, user_bni) As Boolean

    Dim klm(), isi()
    Dim nmTabel As String
    
    nmTabel = "ubni_direct"
    
    klm = Array("norek", "nmpemegang", "user_bni")
    isi = Array(noRek, nmPemegang, user_bni)
    If tbUpdate(nmTabel, klm, isi, cnn, "kodedivisi = '" & kodeDivisi & "'") = True Then
        tbBniDirect_update = True
    Else
        tbBniDirect_update = False
    End If
End Function

Function tbBniDirect_getNorek(kdDivisi As String) As String
    Dim sql As String
    
    sql = "select norek from ubni_direct where kodedivisi = '" & kdDivisi & "'"
    tbBniDirect_getNorek = cari_data1(cnn, sql)
End Function

Function tbBniDirect_get(kdDivisi As String, kolom As String) As String
    Dim sql As String
    
    sql = "select " & kolom & " from ubni_direct where kodedivisi = '" & kdDivisi & "'"
    tbBniDirect_get = cari_data1(cnn, sql)
End Function

Function tbBniDirect_delete(id1 As String) As Boolean

    Dim klm(), isi()
    Dim nmTabel As String
    
    nmTabel = "ubni_direct"
    
    klm = Array("id1")
    isi = Array(id1)
    If tbDelete(nmTabel, klm, isi, cnn) = True Then
        tbBniDirect_delete = True
    Else
        tbBniDirect_delete = False
    End If
End Function

