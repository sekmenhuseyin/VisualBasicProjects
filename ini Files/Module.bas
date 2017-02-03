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
Public Function AyarlarýOku(Bölüm As String, Anahtar As String, Varsayýlan As String) As String
    Dim Deðer As String
    Dim IniFile As String
    Dim FuncLength As Long
    IniFile = App.Path + "\settings.ini"
    Deðer = Space(255)
    FuncLength = GetPrivateProfileString(Bölüm, Anahtar, Varsayýlan, Deðer, 255, IniFile)
    Deðer = Left(Deðer, FuncLength)
    AyarlarýOku = Deðer
End Function
Public Function AyarlarýKaydet(Bölüm As String, Anahtar As String, Deðer As String) As String
    Dim IniFile As String
    Dim FuncLength As Long
    IniFile = App.Path + "\settings.ini"
    FuncLength = WritePrivateProfileString(Bölüm, Anahtar, Deðer, IniFile)
    AyarlarýKaydet = FuncLength
End Function

