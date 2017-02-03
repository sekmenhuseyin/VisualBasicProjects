Attribute VB_Name = "Module"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpDefault As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpString As Any, ByVal _
     lpFileName As String) As Long
Public Function Ayarlar�Oku(B�l�m As String, Anahtar As String, Varsay�lan As String) As String
    Dim De�er As String
    Dim IniFile As String
    Dim FuncLength As Long
    IniFile = App.Path + "\settings.ini"
    De�er = Space(255)
    FuncLength = GetPrivateProfileString(B�l�m, Anahtar, Varsay�lan, De�er, 255, IniFile)
    De�er = Left(De�er, FuncLength)
    Ayarlar�Oku = De�er
End Function
Public Function Ayarlar�Kaydet(B�l�m As String, Anahtar As String, De�er As String) As String
    Dim IniFile As String
    Dim FuncLength As Long
    IniFile = App.Path + "\settings.ini"
    FuncLength = WritePrivateProfileString(B�l�m, Anahtar, De�er, IniFile)
    Ayarlar�Kaydet = FuncLength
End Function

