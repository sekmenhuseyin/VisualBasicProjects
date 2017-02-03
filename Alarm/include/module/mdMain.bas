Attribute VB_Name = "mdMain"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Global option_NO1 As Byte: Global opt1_txt_Music As String: Global opt1_txt_Repeat As Byte: Global opt1_sld_Volume As Byte
Global option_NO2 As Byte: Global opt2_txt_Program As String
Global option_NO3 As Byte: Global opt3_txt_Message As String
Global option_NO4 As Byte: Global opt4_cmb_Shutdown As Byte: Global opt4_opt_Force As Byte: Global opt4_txt_Time As Byte: Global opt4_txt_Shutdown_Msg As String
Global option_NO5 As Byte: Global opt5_opt_Type As Byte
Global Settings_opt1 As Byte: Global Settings_opt2 As Byte: Global Settings_opt3 As Byte: Global Settings_opt4 As Byte: Global Settings_opt5 As Byte
Global ColorNo1 As String: Global ColorNo2 As String: Global ColorNo3 As String: Global ColorNo4 As String: Global Color_Theme As Byte
Global Time_saat As Byte: Global Time_dakka As Byte
Global temp_X As Integer: Global temp_Y As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Global gizlendiMi As Boolean: Global OptionsOn As Boolean: Global SettingsOn As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type AlarmType
    AlarmDate As String: AlarmSaat As String: AlarmDakka As String
    SndOn As Boolean: SndPath As String: SndRpt As Byte: SndVol As Byte
    MsgOn As Boolean: MsgTxt As String
    AppOn As Boolean: AppPath As String
    ShutOn As Boolean: ShutType As Byte: ShutZor As Boolean: ShutTime As Byte: ShutMsg As String
End Type
Global AlarmSettings() As AlarmType: Global AlarmCount As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Main()
    If App.PrevInstance = True Then End   'daha �nceden a��lm��sa bir daha a�ma
    Non_Changable_Telif_Check 'telif haklar� kontrol� yap�l�yor...
    Call Varsay�lanlar 'varsay�lanlara d�n��. alttaki dosyalar�n olmama ihtimaline kar��!
    If Dir(App.Path + "\settings.ini") <> "" Then Call AyarOku 'settings.ini varsa i�eri�i okur, de�i�kenlere aktar
    If Dir(App.Path + "\alarms.data") <> "" Then Call AlarmOku   'alarms.data varsa i�eri�i okur, de�i�kenlere aktar
    frmAlarm.Show
End Sub
Sub The_End()
    If frmAlarm.WindowState = 0 Then temp_X = frmAlarm.Left: temp_Y = frmAlarm.Top
    If OptionsOn = True Then Unload frmOptions
    If SettingsOn = True Then Unload frmSettings
    Call AyarKaydet: Call AlarmKaydet
    End
End Sub
Private Sub Varsay�lanlar()
    'about option1      ses �al
    option_NO1 = 1
    opt1_txt_Music = App.Path + "\include\sound\horoz.wav"
    opt1_txt_Repeat = 1
    opt1_sld_Volume = 100
    'about option2      uygulama �al��r�t
    option_NO2 = 0
    opt2_txt_Program = ""
    'about option3      mesaj
    option_NO3 = 0
    opt3_txt_Message = ""
    'about option4      bilgisayar� kapat
    option_NO4 = 0
    opt4_cmb_Shutdown = 0
    opt4_opt_Force = 0
    opt4_txt_Time = 30
    opt4_txt_Shutdown_Msg = ""
    'about option5      saat ba�lar�
    option_NO5 = 0
    opt5_opt_Type = 1
    'Settings
    Settings_opt1 = 0
    Settings_opt2 = 0
    Settings_opt3 = 0
    Settings_opt4 = 0
    Settings_opt5 = 0
    'G�r�n�m
    ColorNo1 = RGB(212, 214, 186)
    ColorNo2 = RGB(0, 0, 0)
    ColorNo3 = RGB(236, 233, 216)
    ColorNo4 = RGB(0, 0, 0)
    Color_Theme = 1 'Standart
    'alarm time
    Time_dakka = 0
    Time_saat = 0
    'koordinats
    temp_X = 0
    temp_Y = 0
End Sub
Private Sub AlarmOku()

End Sub
Private Sub AlarmKaydet()

End Sub
Private Sub AyarOku()
    'about option1      ses �al
    option_NO1 = ReadStringFromIni("option1", "option_NO1", "1")
    opt1_txt_Music = ReadStringFromIni("option1", "opt1_txt_Music", App.Path + "\include\sound\horoz.wav")
    opt1_txt_Repeat = ReadStringFromIni("option1", "opt1_txt_Repeat", "1")
    opt1_sld_Volume = ReadStringFromIni("option1", "opt1_sld_Volume", "100")
    'about option2      uygulama �al��t�r
    option_NO2 = ReadStringFromIni("option2", "option_NO2", "0")
    opt2_txt_Program = ReadStringFromIni("option2", "opt2_txt_Program", "")
    'about option3      mesaj
    option_NO3 = ReadStringFromIni("option3", "option_NO3", "0")
    opt3_txt_Message = ReadStringFromIni("option3", "opt3_txt_Message", "")
    'about option4      bilgisayar� kapat
    option_NO4 = ReadStringFromIni("option4", "option_NO4", "0")
    opt4_cmb_Shutdown = ReadStringFromIni("option4", "opt4_cmb_Shutdown", "0")
    opt4_opt_Force = ReadStringFromIni("option4", "opt4_opt_Force", "0")
    opt4_txt_Time = ReadStringFromIni("option4", "opt4_txt_Time", "30")
    opt4_txt_Shutdown_Msg = ReadStringFromIni("option4", "opt4_txt_Shutdown_Msg", "")
    'about option5      saat ba�lar�
    option_NO5 = ReadStringFromIni("option5", "option_NO5", "0")
    opt5_opt_Type = ReadStringFromIni("option5", "opt5_opt_Type", "1")
    'Settings
    Settings_opt1 = ReadStringFromIni("Settings", "Settings_opt1", "0")
    Settings_opt2 = ReadStringFromIni("Settings", "Settings_opt2", "0")
    Settings_opt3 = ReadStringFromIni("Settings", "Settings_opt3", "0")
    Settings_opt4 = ReadStringFromIni("Settings", "Settings_opt4", "0")
    Settings_opt5 = ReadStringFromIni("Settings", "Settings_opt5", "0")
    'G�r�n�m
    ColorNo1 = ReadStringFromIni("G�r�n�m", "ColorNo1", "-2147483645")
    ColorNo2 = ReadStringFromIni("G�r�n�m", "ColorNo2", "0")
    ColorNo3 = ReadStringFromIni("G�r�n�m", "ColorNo3", "-2147483633")
    ColorNo4 = ReadStringFromIni("G�r�n�m", "ColorNo4", "0")
    Color_Theme = ReadStringFromIni("G�r�n�m", "Color_Theme", "1")
    'Zaman
    Time_dakka = ReadStringFromIni("Zaman", "Time_dakka", "0")
    Time_saat = ReadStringFromIni("Zaman", "Time_saat", "0")
    'Koordinat
    temp_X = ReadStringFromIni("Koordinat", "temp_X", "0")
    temp_Y = ReadStringFromIni("Koordinat", "temp_Y", "0")
End Sub
Sub AyarKaydet()
    'about option1      ses �al
    WriteStringToIni "option1", "option_NO1", CStr(option_NO1)
    WriteStringToIni "option1", "opt1_txt_Music", CStr(opt1_txt_Music)
    WriteStringToIni "option1", "opt1_txt_Repeat", CStr(opt1_txt_Repeat)
    WriteStringToIni "option1", "opt1_sld_Volume", CStr(opt1_sld_Volume)
    'about option2      uygulama �al��t�r
    WriteStringToIni "option2", "option_NO2", CStr(option_NO2)
    WriteStringToIni "option2", "opt2_txt_Program", CStr(opt2_txt_Program)
    'about option3      mesaj
    WriteStringToIni "option3", "option_NO3", CStr(option_NO3)
    WriteStringToIni "option3", "opt3_txt_Message", CStr(opt3_txt_Message)
    'about option4      bilgisayar� kapat
    WriteStringToIni "option4", "option_NO4", CStr(option_NO4)
    WriteStringToIni "option4", "opt4_cmb_Shutdown", CStr(opt4_cmb_Shutdown)
    WriteStringToIni "option4", "opt4_opt_Force", CStr(opt4_opt_Force)
    WriteStringToIni "option4", "opt4_txt_Time", CStr(opt4_txt_Time)
    WriteStringToIni "option4", "opt4_txt_Shutdown_Msg", CStr(opt4_txt_Shutdown_Msg)
    'about option5      saat ba�lar�
    WriteStringToIni "option5", "option_NO5", CStr(option_NO5)
    WriteStringToIni "option5", "opt5_opt_Type", CStr(opt5_opt_Type)
    'Settings
    WriteStringToIni "Settings", "Settings_opt1", CStr(Settings_opt1)
    WriteStringToIni "Settings", "Settings_opt2", CStr(Settings_opt2)
    WriteStringToIni "Settings", "Settings_opt3", CStr(Settings_opt3)
    WriteStringToIni "Settings", "Settings_opt4", CStr(Settings_opt4)
    WriteStringToIni "Settings", "Settings_opt5", CStr(Settings_opt5)
    'G�r�n�m
    WriteStringToIni "G�r�n�m", "ColorNo1", CStr(ColorNo1)
    WriteStringToIni "G�r�n�m", "ColorNo2", CStr(ColorNo2)
    WriteStringToIni "G�r�n�m", "ColorNo3", CStr(ColorNo3)
    WriteStringToIni "G�r�n�m", "ColorNo4", CStr(ColorNo4)
    WriteStringToIni "G�r�n�m", "Color_Theme", CStr(Color_Theme)
    'Zaman
    WriteStringToIni "Zaman", "Time_dakka", CStr(frmAlarm.dakka.ListIndex)
    WriteStringToIni "Zaman", "Time_saat", CStr(frmAlarm.saat.ListIndex)
    'Koordinat
    WriteStringToIni "Koordinat", "temp_X", CStr(temp_X)
    WriteStringToIni "Koordinat", "temp_Y", CStr(temp_Y)
End Sub
Private Sub Non_Changable_Telif_Check()
    If App.CompanyName = "Sekmenler Tech." Then
        If App.LegalCopyright = "� " + App.CompanyName Then Exit Sub
    End If
    MsgBox "Uygulaman�n telif haklar� de�i�tirilmi�." + Chr(13) + Chr(10) + "L�tfen uygulamay� tekrar kurun."
    End
End Sub
Private Function ReadStringFromIni(B�l�m As String, Anahtar As String, Varsay�lan As String) As String
    Dim De�er As String
    Dim IniFile As String
    Dim FuncLength As Long
    IniFile = App.Path + "\settings.ini"
    De�er = Space(255)
    FuncLength = GetPrivateProfileString(B�l�m, Anahtar, Varsay�lan, De�er, 255, IniFile)
    De�er = Left(De�er, FuncLength)
    ReadStringFromIni = De�er
End Function
Private Function WriteStringToIni(B�l�m As String, Anahtar As String, De�er As String) As String
    Dim IniFile As String
    Dim FuncLength As Long
    IniFile = App.Path + "\settings.ini"
    FuncLength = WritePrivateProfileString(B�l�m, Anahtar, De�er, IniFile)
    WriteStringToIni = FuncLength
End Function
