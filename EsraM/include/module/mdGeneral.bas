Attribute VB_Name = "xmdGeneral"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''ini dosyas�ndan bilgi okuyup yazar''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Global rnk_Yaz�_Arka As String: Global rnk_Frm_Arka As String: Global rnk_Frm_�n As String: Global rnk_Yaz�_�n As String: Global rnk_Btn_Arka As String: Global rnk_Btn_�n As String
Global ResimNO As Integer: Global Tema_Yeri As String: Global Tema_Ad� As String: Global cmbindex As Integer
Private Type ThemePros: TemaAd As String: TemaDizin As String: End Type: Global Themes() As ThemePros
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
Sub Main()
    If App.PrevInstance = True Then End 'daha �nceden a��lm��sa bir daha a�ma
    Call Dosya_kontrol 'Gerekli dosyalar kontrol ediliyor.
    Call Non_Changable_Telif_Check 'telif haklar� kontrol� yap�l�yor...
    Call RegisterTheEsraM 'register the controls
    Call CheckLogging 'log kontrol�: d�zg�n kapnamam��sa crach dosyas�na yazar.
    Call Varsay�lanlaraD�n: Call Ayarlar�Oku 'de�i�kenler tan�mlan�r. settings.ini varsa okunur
    Write #7, Day(Date) & "." & Month(Date) & "." & Year(Date), "mdGeneral", "Main", "Successful" 'logging
    frmSplash.Show 'SPLASH
End Sub
Sub The_End()
    Dim Form As Form: For Each Form In Forms: Unload Form: Set Form = Nothing: Next Form 'unloading forms
    Write #7, Day(Date) & "." & Month(Date) & "." & Year(Date), "mdGeneral", "The_End", "Successful" 'logging
    Call Ayarlar�Kaydet: Close
    If Dir(App.path + "\temp.txt") <> "" Then Kill App.path + "\temp.txt"
    If Dir(App.path + "\temp2.txt") <> "" Then Kill App.path + "\temp2.txt"
    If Dir(App.path + "\temp3.txt") <> "" Then Kill App.path + "\temp3.txt"
    Kill App.path & "\" & App.EXEName & ".exe.log" 'nas�l olsa d�zg�n kapand� art�k...
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
Private Sub Varsay�lanlaraD�n()
    rnk_Frm_Arka = "12648384": rnk_Frm_�n = "0"
    rnk_Yaz�_Arka = "16777215": rnk_Yaz�_�n = "0"
    rnk_Btn_Arka = "12648384": rnk_Btn_�n = "0"
    Tema_Ad� = "EsraM Standart": cmbindex = 0: ResimNO = 0
End Sub
Private Sub Ayarlar�Oku()
    ResimNO = ReadStringFromIni("Resim", "ResimNO", CStr(ResimNO))
    Tema_Ad� = ReadStringFromIni("Uygulama", "Tema_Ad�", CStr(Tema_Ad�))
    cmbindex = ReadStringFromIni("Uygulama", "cmbindex", CStr(cmbindex))
    rnk_Frm_Arka = ReadStringFromIni("G�r�n�m", "rnk_frm_arka", CStr(rnk_Frm_Arka))
    rnk_Frm_�n = ReadStringFromIni("G�r�n�m", "rnk_frm_�n", CStr(rnk_Frm_�n))
    rnk_Yaz�_Arka = ReadStringFromIni("G�r�n�m", "rnk_yaz�_arka", CStr(rnk_Yaz�_Arka))
    rnk_Yaz�_�n = ReadStringFromIni("G�r�n�m", "rnk_yaz�_�n", CStr(rnk_Yaz�_�n))
    rnk_Btn_Arka = ReadStringFromIni("G�r�n�m", "rnk_btn_arka", CStr(rnk_Btn_Arka))
    rnk_Btn_�n = ReadStringFromIni("G�r�n�m", "rnk_btn_�n", CStr(rnk_Btn_�n))
End Sub
Private Sub Ayarlar�Kaydet()
    WriteStringToIni "Resim", "ResimNO", CStr(ResimNO)
    WriteStringToIni "Uygulama", "Tema_Ad�", CStr(Tema_Ad�)
    WriteStringToIni "Uygulama", "OpenClose", "Closed"
    WriteStringToIni "Uygulama", "cmbindex", CStr(cmbindex)
    WriteStringToIni "G�r�n�m", "rnk_frm_arka", CStr(rnk_Frm_Arka)
    WriteStringToIni "G�r�n�m", "rnk_frm_�n", CStr(rnk_Frm_�n)
    WriteStringToIni "G�r�n�m", "rnk_yaz�_arka", CStr(rnk_Yaz�_Arka)
    WriteStringToIni "G�r�n�m", "rnk_yaz�_�n", CStr(rnk_Yaz�_�n)
    WriteStringToIni "G�r�n�m", "rnk_btn_arka", CStr(rnk_Btn_Arka)
    WriteStringToIni "G�r�n�m", "rnk_btn_�n", CStr(rnk_Btn_�n)
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
Sub Dosya_kontrol()
    On Local Error Resume Next
    Dim Dosyalar�n_Durumu_iyimi As Boolean
    Dosyalar�n_Durumu_iyimi = True
    'i�te �nemli dosya ve klas�rler
    If Dir(App.path + "\Style", vbDirectory) = "" Then MkDir App.path + "\Style"
    If Dir(App.path + "\Style\EsraM Standart", vbDirectory) = "" Then MkDir App.path + "\Style\EsraM Standart"
    If Dir(App.path + "\Style\EsraM Standart\style.ini") = "" Then
        WriteStringToIni "Theme", "rnk_yaz�_arka", "16777215", App.path + "\Style\EsraM Standart"
        WriteStringToIni "Theme", "rnk_frm_arka", "12648384", App.path + "\Style\EsraM Standart"
        WriteStringToIni "Theme", "rnk_yaz�_�n", "0", App.path + "\Style\EsraM Standart"
        WriteStringToIni "Theme", "rnk_btn_arka", "12648384", App.path + "\Style\EsraM Standart"
        WriteStringToIni "Theme", "rnk_btn_�n", "0", App.path + "\Style\EsraM Standart"
    End If
    If Dir(App.path + "\Pictures\sample.qaz") = "" Then Dosyalar�n_Durumu_iyimi = False
    If Dir(App.path + "\include\data\Data.mdb") = "" Then Dosyalar�n_Durumu_iyimi = False
    If Dir(App.path + "\include\data\credits.avi") = "" Then Dosyalar�n_Durumu_iyimi = False
    If Dir(App.path + "\Resimci.exe") = "" Then Dosyalar�n_Durumu_iyimi = False
    '�nemli dosyalar eksikse �al��ma!
    If Dosyalar�n_Durumu_iyimi = False Then MsgBox "Program�n dosyalar� hasar g�rm��. L�tfen program� tekrar y�kleyin !", vbCritical + vbApplicationModal, "Hatal� Y�kleme": End
End Sub
Private Sub Non_Changable_Telif_Check()
    If App.CompanyName = "Sekmenler Tech." Then
        If App.LegalCopyright = "� " + App.CompanyName Then Exit Sub
    End If
    MsgBox "Uygulaman�n telif haklar� de�i�tirilmi�." + Chr(13) + Chr(10) + "L�tfen uygulamay� tekrar kurun."
    End
End Sub
Private Sub RegisterTheEsraM()
    'windows\system klas�r� bulunuyor
    'e�er register yap�lmam��sa orada bir .reg dosyas� olu�turulacak.
    'e�er bu dosya varsa register yap�ld���ndan rahatl�kla bu yugulama �al��acak!!!
    On Local Error Resume Next: Dim tempSTR As String: Dim i As Integer: Randomize
    If Dir(Environ("windir") & "\system\EsraM.rst") = "" Then
        Shell App.path + "\include\register\register.bat", vbMinimizedNoFocus
        '*.rst dosyas�na rastgele bir�eyler yaz�yoruz ki g�renler bir �ey yaz�yor sans�n.
        For i = 0 To 32500: tempSTR = tempSTR & Chr(Int(Rnd * 200) + 50): Next i
        Open Environ("windir") & "\system\EsraM.rst" For Output As #1: Print #1, tempSTR: Close #1
    End If
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*

Public Function ReadStringFromIni(B�l�m As String, Anahtar As String, Varsay�lan As String, Optional Style_Yolu As String) As String
    Dim De�er As String: Dim IniFile As String: Dim FuncLength As Long
    If Style_Yolu = "" Then IniFile = App.path + "\settings.ini" Else IniFile = Style_Yolu + "\style.ini"
    De�er = Space(255)
    FuncLength = GetPrivateProfileString(B�l�m, Anahtar, Varsay�lan, De�er, 255, IniFile)
    De�er = Left(De�er, FuncLength)
    ReadStringFromIni = De�er
End Function
Private Function WriteStringToIni(B�l�m As String, Anahtar As String, De�er As String, Optional Style_Yolu As String) As String
    Dim IniFile As String: Dim FuncLength As Long
    If Style_Yolu = "" Then IniFile = App.path + "\settings.ini" Else IniFile = Style_Yolu + "\style.ini"
    FuncLength = WritePrivateProfileString(B�l�m, Anahtar, De�er, IniFile)
    WriteStringToIni = FuncLength
End Function
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*

