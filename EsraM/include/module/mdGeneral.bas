Attribute VB_Name = "xmdGeneral"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''ini dosyasýndan bilgi okuyup yazar''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Global rnk_Yazý_Arka As String: Global rnk_Frm_Arka As String: Global rnk_Frm_Ön As String: Global rnk_Yazý_Ön As String: Global rnk_Btn_Arka As String: Global rnk_Btn_Ön As String
Global ResimNO As Integer: Global Tema_Yeri As String: Global Tema_Adý As String: Global cmbindex As Integer
Private Type ThemePros: TemaAd As String: TemaDizin As String: End Type: Global Themes() As ThemePros
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
Sub Main()
    If App.PrevInstance = True Then End 'daha önceden açýlmýþsa bir daha açma
    Call Dosya_kontrol 'Gerekli dosyalar kontrol ediliyor.
    Call Non_Changable_Telif_Check 'telif haklarý kontrolü yapýlýyor...
    Call RegisterTheEsraM 'register the controls
    Call CheckLogging 'log kontrolü: düzgün kapnamamýþsa crach dosyasýna yazar.
    Call VarsayýlanlaraDön: Call AyarlarýOku 'deðiþkenler tanýmlanýr. settings.ini varsa okunur
    Write #7, Day(Date) & "." & Month(Date) & "." & Year(Date), "mdGeneral", "Main", "Successful" 'logging
    frmSplash.Show 'SPLASH
End Sub
Sub The_End()
    Dim Form As Form: For Each Form In Forms: Unload Form: Set Form = Nothing: Next Form 'unloading forms
    Write #7, Day(Date) & "." & Month(Date) & "." & Year(Date), "mdGeneral", "The_End", "Successful" 'logging
    Call AyarlarýKaydet: Close
    If Dir(App.path + "\temp.txt") <> "" Then Kill App.path + "\temp.txt"
    If Dir(App.path + "\temp2.txt") <> "" Then Kill App.path + "\temp2.txt"
    If Dir(App.path + "\temp3.txt") <> "" Then Kill App.path + "\temp3.txt"
    Kill App.path & "\" & App.EXEName & ".exe.log" 'nasýl olsa düzgün kapandý artýk...
    End
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*

Private Sub CheckLogging()
    On Local Error Resume Next
    Dim OpenClose, TimeOfSub, NameofForm, TypeOfSub, Description As String
    OpenClose = ReadStringFromIni("Uygulama", "OpenClose", "Closed")
    If OpenClose = "Opened" And Dir(App.path & "\" & App.EXEName & ".exe.log") <> "" Then
        Open App.path & "\" & App.EXEName & ".exe.log" For Input As #7
        Open App.path & "\crash." & App.EXEName & ".exe.log" For Append As #9
            Do While EOF(7) = False
                TimeOfSub = "": NameofForm = "": TypeOfSub = "": Description = ""
                Input #7, TimeOfSub, NameofForm, TypeOfSub, Description
                Write #9, TimeOfSub, NameofForm, TypeOfSub, Description
            Loop
        Write #9, String(50, "-"): Write #9, String(50, "-"): Write #9, ""
        Close #9: Close #7
    End If
    WriteStringToIni "Uygulama", "OpenClose", "Opened"
    Open App.path & "\" & App.EXEName & ".exe.log" For Output As #7
End Sub
Private Sub VarsayýlanlaraDön()
    rnk_Frm_Arka = "12648384": rnk_Frm_Ön = "0"
    rnk_Yazý_Arka = "16777215": rnk_Yazý_Ön = "0"
    rnk_Btn_Arka = "12648384": rnk_Btn_Ön = "0"
    Tema_Adý = "EsraM Standart": cmbindex = 0: ResimNO = 0
End Sub
Private Sub AyarlarýOku()
    ResimNO = ReadStringFromIni("Resim", "ResimNO", CStr(ResimNO))
    Tema_Adý = ReadStringFromIni("Uygulama", "Tema_Adý", CStr(Tema_Adý))
    cmbindex = ReadStringFromIni("Uygulama", "cmbindex", CStr(cmbindex))
    rnk_Frm_Arka = ReadStringFromIni("Görünüm", "rnk_frm_arka", CStr(rnk_Frm_Arka))
    rnk_Frm_Ön = ReadStringFromIni("Görünüm", "rnk_frm_ön", CStr(rnk_Frm_Ön))
    rnk_Yazý_Arka = ReadStringFromIni("Görünüm", "rnk_yazý_arka", CStr(rnk_Yazý_Arka))
    rnk_Yazý_Ön = ReadStringFromIni("Görünüm", "rnk_yazý_ön", CStr(rnk_Yazý_Ön))
    rnk_Btn_Arka = ReadStringFromIni("Görünüm", "rnk_btn_arka", CStr(rnk_Btn_Arka))
    rnk_Btn_Ön = ReadStringFromIni("Görünüm", "rnk_btn_ön", CStr(rnk_Btn_Ön))
End Sub
Private Sub AyarlarýKaydet()
    WriteStringToIni "Resim", "ResimNO", CStr(ResimNO)
    WriteStringToIni "Uygulama", "Tema_Adý", CStr(Tema_Adý)
    WriteStringToIni "Uygulama", "OpenClose", "Closed"
    WriteStringToIni "Uygulama", "cmbindex", CStr(cmbindex)
    WriteStringToIni "Görünüm", "rnk_frm_arka", CStr(rnk_Frm_Arka)
    WriteStringToIni "Görünüm", "rnk_frm_ön", CStr(rnk_Frm_Ön)
    WriteStringToIni "Görünüm", "rnk_yazý_arka", CStr(rnk_Yazý_Arka)
    WriteStringToIni "Görünüm", "rnk_yazý_ön", CStr(rnk_Yazý_Ön)
    WriteStringToIni "Görünüm", "rnk_btn_arka", CStr(rnk_Btn_Arka)
    WriteStringToIni "Görünüm", "rnk_btn_ön", CStr(rnk_Btn_Ön)
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
Sub Dosya_kontrol()
    On Local Error Resume Next
    Dim Dosyalarýn_Durumu_iyimi As Boolean
    Dosyalarýn_Durumu_iyimi = True
    'iþte önemli dosya ve klasörler
    If Dir(App.path + "\Style", vbDirectory) = "" Then MkDir App.path + "\Style"
    If Dir(App.path + "\Style\EsraM Standart", vbDirectory) = "" Then MkDir App.path + "\Style\EsraM Standart"
    If Dir(App.path + "\Style\EsraM Standart\style.ini") = "" Then
        WriteStringToIni "Theme", "rnk_yazý_arka", "16777215", App.path + "\Style\EsraM Standart"
        WriteStringToIni "Theme", "rnk_frm_arka", "12648384", App.path + "\Style\EsraM Standart"
        WriteStringToIni "Theme", "rnk_yazý_ön", "0", App.path + "\Style\EsraM Standart"
        WriteStringToIni "Theme", "rnk_btn_arka", "12648384", App.path + "\Style\EsraM Standart"
        WriteStringToIni "Theme", "rnk_btn_ön", "0", App.path + "\Style\EsraM Standart"
    End If
    If Dir(App.path + "\Pictures\sample.qaz") = "" Then Dosyalarýn_Durumu_iyimi = False
    If Dir(App.path + "\include\data\Data.mdb") = "" Then Dosyalarýn_Durumu_iyimi = False
    If Dir(App.path + "\include\data\credits.avi") = "" Then Dosyalarýn_Durumu_iyimi = False
    If Dir(App.path + "\Resimci.exe") = "" Then Dosyalarýn_Durumu_iyimi = False
    'önemli dosyalar eksikse çalýþma!
    If Dosyalarýn_Durumu_iyimi = False Then MsgBox "Programýn dosyalarý hasar görmüþ. Lütfen programý tekrar yükleyin !", vbCritical + vbApplicationModal, "Hatalý Yükleme": End
End Sub
Private Sub Non_Changable_Telif_Check()
    If App.CompanyName = "Sekmenler Tech." Then
        If App.LegalCopyright = "© " + App.CompanyName Then Exit Sub
    End If
    MsgBox "Uygulamanýn telif haklarý deðiþtirilmiþ." + Chr(13) + Chr(10) + "Lütfen uygulamayý tekrar kurun."
    End
End Sub
Private Sub RegisterTheEsraM()
    'windows\system klasörü bulunuyor
    'eðer register yapýlmamýþsa orada bir .reg dosyasý oluþturulacak.
    'eðer bu dosya varsa register yapýldýðýndan rahatlýkla bu yugulama çalýþacak!!!
    On Local Error Resume Next: Dim tempSTR As String: Dim i As Integer: Randomize
    If Dir(Environ("windir") & "\system\EsraM.rst") = "" Then
        Shell App.path + "\include\register\register.bat", vbMinimizedNoFocus
        '*.rst dosyasýna rastgele birþeyler yazýyoruz ki görenler bir þey yazýyor sansýn.
        For i = 0 To 32500: tempSTR = tempSTR & Chr(Int(Rnd * 200) + 50): Next i
        Open Environ("windir") & "\system\EsraM.rst" For Output As #1: Print #1, tempSTR: Close #1
    End If
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*

Public Function ReadStringFromIni(Bölüm As String, Anahtar As String, Varsayýlan As String, Optional Style_Yolu As String) As String
    Dim Deðer As String: Dim IniFile As String: Dim FuncLength As Long
    If Style_Yolu = "" Then IniFile = App.path + "\settings.ini" Else IniFile = Style_Yolu + "\style.ini"
    Deðer = Space(255)
    FuncLength = GetPrivateProfileString(Bölüm, Anahtar, Varsayýlan, Deðer, 255, IniFile)
    Deðer = Left(Deðer, FuncLength)
    ReadStringFromIni = Deðer
End Function
Private Function WriteStringToIni(Bölüm As String, Anahtar As String, Deðer As String, Optional Style_Yolu As String) As String
    Dim IniFile As String: Dim FuncLength As Long
    If Style_Yolu = "" Then IniFile = App.path + "\settings.ini" Else IniFile = Style_Yolu + "\style.ini"
    FuncLength = WritePrivateProfileString(Bölüm, Anahtar, Deðer, IniFile)
    WriteStringToIni = FuncLength
End Function
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*

