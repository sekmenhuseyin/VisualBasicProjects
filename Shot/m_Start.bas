Attribute VB_Name = "m_Start"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''ini dosyas�ndan bilgi okur
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
''''''''''''''''''''''''''''''''''''''''''''''ini dosyas�na bilgi yazar
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
''''''''''''''''''''''''''''''''''''''''''''''windows klas�r�n� ��renir
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Global Variables
Global Best_Name(9) As String: Global Best_Score(9) As Integer
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
Sub Main()
    'daha �nceden a��lm��sa bir daha a�ma
    If App.PrevInstance = True Then End
    'register kontrol
    RegisterTheAPP
    'telif haklar� kontrol� yap�l�yor...
    Non_Changable_Telif_Check
    'ilk �nce de�i�kenler tan�mlan�r. daha sonra settings.ini okunur
    Varsay�lanlaraD�n
    Ayarlar�Oku
    'SPLASH
    f_Main.Show
End Sub
Sub The_End()
    Ayarlar�Kaydet
    Unload f_Main
    Unload f_Scores
    Unload f_Game
    End
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*

Private Sub Varsay�lanlaraD�n()
    Dim i As Byte: Dim tmp As Integer
    tmp = 27
    For i = 0 To 9
        Best_Name(i) = Chr(i + 65) + Chr(97 + i) + Chr(97 + i)
        Best_Score(9 - i) = tmp + (2 * i + 3)
        tmp = Best_Score(9 - i)
    Next i
End Sub
Private Sub Ayarlar�Oku()
    Dim i As Byte
    For i = 0 To 9
        Best_Name(i) = ReadStringFromIni("Name", "Best_Name_" + CStr(i), Best_Name(i))
        Best_Score(i) = ReadStringFromIni("Score", "Best_Score_" + CStr(i), CStr(Best_Score(i)))
    Next i
End Sub
Public Sub Ayarlar�Kaydet()
    Dim tmp As Long: Dim i As Byte
    For i = 0 To 9
        tmp = WriteStringToIni("Name", "Best_Name_" + CStr(i), Best_Name(i))
        tmp = WriteStringToIni("Score", "Best_Score_" + CStr(i), CStr(Best_Score(i)))
    Next i
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
Private Sub Non_Changable_Telif_Check()
    If App.CompanyName = "Sekmenler Tech." Then
        If App.LegalCopyright = "� " + App.CompanyName Then Exit Sub
    End If
    MsgBox "Uygulaman�n telif haklar� de�i�tirilmi�." + Chr(13) + Chr(10) + "L�tfen uygulamay� tekrar kurun."
    End
End Sub
Private Sub RegisterTheAPP()
    'windows\system klas�r� bulunuyor
    'e�er register yap�lmam��sa orada bir .reg dosyas� olu�turulacak.
    'e�er bu dosya varsa register yap�ld���ndan rahatl�kla bu yugulama �al��acak!!!
    Dim Yol As String: Dim Uzunluk As Integer
    Dim tempSTR As String: Dim i As Integer
    Randomize
    Yol = Space(255)
    Uzunluk = GetWindowsDirectory(Yol, Len(Yol))
    If Dir(Left(Yol, Uzunluk) + "\system\Shot.rst") = "" Then
        'Shell App.Path + "\register\register.bat", vbMinimizedNoFocus
        '*.rst dosyas�na rastgele bir�eyler yaz�yoruz ki g�renler bir �ey yaz�yor sans�n.
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
    'ilk �nce ilk harf hari� t�m harfleri k���lt�yoruz.
    For i = 2 To Len(OldString)
        If Mid(OldString, i, 1) = "�" Then
            Mid(OldString, i, 1) = "i"
        ElseIf Mid(OldString, i, 1) = "I" Then
            Mid(OldString, i, 1) = "�"
        Else
            Mid(OldString, i, 1) = LCase(Mid(OldString, i, 1))
        End If
    Next i
    i = 0
basad�n:
    i = i + 1
    'daha sonra sadece ilk harfi b�y�t�yoruz.
    If Mid(OldString, i, 1) = "i" Then
        Mid(OldString, i, 1) = "�"
    ElseIf Mid(OldString, i, 1) = "�" Then
        Mid(OldString, i, 1) = "I"
    Else
        Mid(OldString, i, 1) = UCase(Mid(OldString, i, 1))
    End If
    'ba�ka kelime varsa onlar�n da ilk harflerini b�y�tecez
    i = i + 1: i = InStr(i, OldString, " ")
    If i <> 0 Then GoTo basad�n
    'en son olarak da yukar�da buldu�umuz iki sonucu birle�tirip geriye o stringi d�nd�r�yoruz.
    UpperCaseFirstLetter = OldString
End Function
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*

Public Function ReadStringFromIni(B�l�m As String, Anahtar As String, Varsay�lan As String, Optional Style_Yolu As String) As String
    Dim De�er As String: Dim IniFile As String: Dim FuncLength As Long
    If Style_Yolu = "" Then IniFile = App.Path + "\settings.ini" Else IniFile = Style_Yolu + "\Style.ini"
    De�er = Space(255)
    FuncLength = GetPrivateProfileString(B�l�m, Anahtar, Varsay�lan, De�er, 255, IniFile)
    De�er = Left(De�er, FuncLength)
    ReadStringFromIni = De�er
End Function
Public Function WriteStringToIni(B�l�m As String, Anahtar As String, De�er As String) As String
    Dim IniFile As String: Dim FuncLength As Long
    IniFile = App.Path + "\settings.ini"
    FuncLength = WritePrivateProfileString(B�l�m, Anahtar, De�er, IniFile)
    WriteStringToIni = FuncLength
End Function
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*



