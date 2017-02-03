Attribute VB_Name = "md_ini"
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
Public Function AyarOku(Bölüm As String, Anahtar As String, Varsayýlan As String) As String
    Dim Deðer As String
    Dim IniFile As String
    Dim FuncLength As Long
    IniFile = App.Path + "\settings.ini"
    Deðer = Space(255)
    FuncLength = GetPrivateProfileString(Bölüm, Anahtar, Varsayýlan, Deðer, 255, IniFile)
    Deðer = Left(Deðer, FuncLength)
    AyarOku = Deðer
End Function
Public Function AyarKaydet(Bölüm As String, Anahtar As String, Deðer As String) As String
    Dim IniFile As String
    Dim FuncLength As Long
    IniFile = App.Path + "\settings.ini"
    FuncLength = WritePrivateProfileString(Bölüm, Anahtar, Deðer, IniFile)
    AyarKaydet = FuncLength
End Function

