Attribute VB_Name = "mdAyar"
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
Sub alarm_SesÇal()                  '''''''''''''''''''''''''''''''''''''''''''''option_NO1
    If Dir(opt1_txt_Music) = "" Then
        MsgBox "Alarm çalýyor !" + Chr(13) + Chr(10) + "Çalýnacak müzik bulunamadý"
        Exit Sub
    End If
    frmAlarm.WMP.URL = opt1_txt_Music
    frmAlarm.WMP.settings.playCount = opt1_txt_Repeat
    frmAlarm.WMP.settings.volume = opt1_sld_Volume
    frmAlarm.WMP.Controls.play
End Sub
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
Sub alarm_UygulamaÇalýþtýr()        '''''''''''''''''''''''''''''''''''''''''''''option_NO2
    On Error Resume Next
    Shell (opt2_txt_Program)
End Sub
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
Sub alarm_MesajVer()                '''''''''''''''''''''''''''''''''''''''''''''option_NO3
    MsgBox opt3_txt_Message
End Sub
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
Sub alarm_BilgisayarýKapat()        '''''''''''''''''''''''''''''''''''''''''''''option_NO4
    Dim temp As String
    Select Case opt4_cmb_Shutdown
        Case "0"    'bilgisayarý kapat
            temp = "shutdown -s"
        Case "1"    'yeniden baþlat
            temp = "shutdown -r"
        Case "2"    'oturumu kapat
            temp = "shutdown -l"
    End Select
    If opt4_opt_Force = 1 Then temp = temp + " -f"
    temp = temp + " -t " + CStr(opt4_txt_Time)
    temp = temp + " -c """ + CStr(opt4_txt_Shutdown_Msg) + """"
    Shell temp
End Sub
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
Sub alarm_SaatBaþý()                '''''''''''''''''''''''''''''''''''''''''''''option_NO5
    If CStr(Right(Time, 5)) = "00:00" Then
        Dim Saat_Baþý_Çalma_Sesi As String
        Saat_Baþý_Çalma_Sesi = Saat_Baþý_Çalma_Sesini_Bul
        If Saat_Baþý_Çalma_Sesi = 0 Then Exit Sub 'saat baþlarýndaki çalma sesini ayarlar.
        If opt5_opt_Type = 0 Then         'Klasik
            frmAlarm.WMP.settings.playCount = 1
            frmAlarm.WMP.URL = Saat_Baþý_Çalma_Sesi
            frmAlarm.WMP.settings.volume = 100
            frmAlarm.WMP.Controls.play
        ElseIf opt5_opt_Type = 1 Then     'Geleneksel
            If Val(Left(Time, 2)) = 0 Then frmAlarm.WMP.settings.playCount = 24 Else frmAlarm.WMP.settings.playCount = Val(Left(Time, 2))
            frmAlarm.WMP.URL = Saat_Baþý_Çalma_Sesi
            frmAlarm.WMP.settings.volume = 100
            frmAlarm.WMP.Controls.play
        ElseIf opt5_opt_Type = 2 Then     'Modern
            Beep
        End If
    End If
End Sub
Function Saat_Baþý_Çalma_Sesini_Bul()
    If Dir(App.Path + "\include\sound\guguk.wav") <> "" Then
        Saat_Baþý_Çalma_Sesini_Bul = App.Path + "\include\sound\guguk.wav"
    Else
        If Dir(Environ("windir") & "\media\ding.wav") <> "" Then
            Saat_Baþý_Çalma_Sesini_Bul = Environ("windir") & "\media\ding.wav"
        Else
            Saat_Baþý_Çalma_Sesini_Bul = 0
        End If
    End If
End Function
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
''''''''''''''''''''''''''''''''''''''''''Renk Temalarý''''''''''''''''''''''''''''''''''''
Sub RenkTemalarý(Theme_Name As String)
    Select Case Theme_Name
        Case "Standart"
            ColorNo1 = RGB(36, 102, 126)
            ColorNo2 = RGB(224, 224, 224)
            ColorNo3 = RGB(224, 224, 224)
            ColorNo4 = RGB(36, 102, 126)
        Case "Mavimsi"
            ColorNo1 = RGB(0, 73, 147)
            ColorNo2 = RGB(255, 255, 255)
            ColorNo3 = RGB(0, 114, 168)
            ColorNo4 = RGB(0, 0, 0)
        Case "Kömür Karasý"
            ColorNo1 = RGB(0, 0, 0)
            ColorNo2 = RGB(255, 255, 255)
            ColorNo3 = RGB(95, 95, 95)
            ColorNo4 = RGB(0, 0, 0)
        Case "Windows XP Temasý"
            ColorNo1 = -2147483645
            ColorNo2 = -2147483640
            ColorNo3 = -2147483633
            ColorNo4 = -2147483640
    End Select
End Sub
''''''''''''''''''''''''''''''''''''''''''Renk Temalarý''''''''''''''''''''''''''''''''''''
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-

