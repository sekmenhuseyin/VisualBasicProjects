Attribute VB_Name = "m_Start"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''ini dosyasýndan bilgi okur
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
''''''''''''''''''''''''''''''''''''''''''''''ini dosyasýna bilgi yazar
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
''''''''''''''''''''''''''''''''''''''''''''''windows klasörünü öðrenir
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Global Variables
Global Best_Name(9) As String: Global Best_Score(9) As Integer
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
Sub Main()
    'daha önceden açýlmýþsa bir daha açma
    If App.PrevInstance = True Then End
    'register kontrol
    RegisterTheAPP
    'telif haklarý kontrolü yapýlýyor...
    Non_Changable_Telif_Check
    'ilk önce deðiþkenler tanýmlanýr. daha sonra settings.ini okunur
    VarsayýlanlaraDön
    AyarlarýOku
    'SPLASH
    f_Main.Show
End Sub
Sub The_End()
    AyarlarýKaydet
    Unload f_Main
    Unload f_Scores
    Unload f_Game
    End
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*

Private Sub VarsayýlanlaraDön()
    Dim i As Byte: Dim tmp As Integer
    tmp = 27
    For i = 0 To 9
        Best_Name(i) = Chr(i + 65) + Chr(97 + i) + Chr(97 + i)
        Best_Score(9 - i) = tmp + (2 * i + 3)
        tmp = Best_Score(9 - i)
    Next i
End Sub
Private Sub AyarlarýOku()
    Dim i As Byte
    For i = 0 To 9
        Best_Name(i) = ReadStringFromIni("Name", "Best_Name_" + CStr(i), Best_Name(i))
        Best_Score(i) = ReadStringFromIni("Score", "Best_Score_" + CStr(i), CStr(Best_Score(i)))
    Next i
End Sub
Public Sub AyarlarýKaydet()
    Dim tmp As Long: Dim i As Byte
    For i = 0 To 9
        tmp = WriteStringToIni("Name", "Best_Name_" + CStr(i), Best_Name(i))
        tmp = WriteStringToIni("Score", "Best_Score_" + CStr(i), CStr(Best_Score(i)))
    Next i
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
Private Sub Non_Changable_Telif_Check()
    If App.CompanyName = "Sekmenler Tech." Then
        If App.LegalCopyright = "© " + App.CompanyName Then Exit Sub
    End If
    MsgBox "Uygulamanýn telif haklarý deðiþtirilmiþ." + Chr(13) + Chr(10) + "Lütfen uygulamayý tekrar kurun."
    End
End Sub
Private Sub RegisterTheAPP()
    'windows\system klasörü bulunuyor
    'eðer register yapýlmamýþsa orada bir .reg dosyasý oluþturulacak.
    'eðer bu dosya varsa register yapýldýðýndan rahatlýkla bu yugulama çalýþacak!!!
    Dim Yol As String: Dim Uzunluk As Integer
    Dim tempSTR As String: Dim i As Integer
    Randomize
    Yol = Space(255)
    Uzunluk = GetWindowsDirectory(Yol, Len(Yol))
    If Dir(Left(Yol, Uzunluk) + "\system\Shot.rst") = "" Then
        'Shell App.Path + "\register\register.bat", vbMinimizedNoFocus
        '*.rst dosyasýna rastgele birþeyler yazýyoruz ki görenler bir þey yazýyor sansýn.
        For i = 0 To 32500
            tempSTR = tempSTR & Chr(Int(Rnd * 200) + 50)
        Next i
        Open Left(Yol, Uzunluk) + "\system\Shot.rst" For Output As #1
            Print #1, tempSTR
        Close #1
    End If
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
Public Function UpperCaseFirstLetter(ByVal OldString As String) As String
    If Trim(OldString) = "" Then Exit Function: Dim i As Integer
    'ilk önce ilk harf hariç tüm harfleri küçültüyoruz.
    For i = 2 To Len(OldString)
        If Mid(OldString, i, 1) = "Ý" Then
            Mid(OldString, i, 1) = "i"
        ElseIf Mid(OldString, i, 1) = "I" Then
            Mid(OldString, i, 1) = "ý"
        Else
            Mid(OldString, i, 1) = LCase(Mid(OldString, i, 1))
        End If
    Next i
    i = 0
basadön:
    i = i + 1
    'daha sonra sadece ilk harfi büyütüyoruz.
    If Mid(OldString, i, 1) = "i" Then
        Mid(OldString, i, 1) = "Ý"
    ElseIf Mid(OldString, i, 1) = "ý" Then
        Mid(OldString, i, 1) = "I"
    Else
        Mid(OldString, i, 1) = UCase(Mid(OldString, i, 1))
    End If
    'baþka kelime varsa onlarýn da ilk harflerini büyütecez
    i = i + 1: i = InStr(i, OldString, " ")
    If i <> 0 Then GoTo basadön
    'en son olarak da yukarýda bulduðumuz iki sonucu birleþtirip geriye o stringi döndürüyoruz.
    UpperCaseFirstLetter = OldString
End Function
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*

Public Function ReadStringFromIni(Bölüm As String, Anahtar As String, Varsayýlan As String, Optional Style_Yolu As String) As String
    Dim Deðer As String: Dim IniFile As String: Dim FuncLength As Long
    If Style_Yolu = "" Then IniFile = App.Path + "\settings.ini" Else IniFile = Style_Yolu + "\Style.ini"
    Deðer = Space(255)
    FuncLength = GetPrivateProfileString(Bölüm, Anahtar, Varsayýlan, Deðer, 255, IniFile)
    Deðer = Left(Deðer, FuncLength)
    ReadStringFromIni = Deðer
End Function
Public Function WriteStringToIni(Bölüm As String, Anahtar As String, Deðer As String) As String
    Dim IniFile As String: Dim FuncLength As Long
    IniFile = App.Path + "\settings.ini"
    FuncLength = WritePrivateProfileString(Bölüm, Anahtar, Deðer, IniFile)
    WriteStringToIni = FuncLength
End Function
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*



