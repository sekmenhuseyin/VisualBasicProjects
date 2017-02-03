Attribute VB_Name = "md_Main"
Option Explicit
'deðiþkenler
Global Konum(2) As String: Global Kenar(1) As String '[Konum]
Global Ayar(6) As String '[Uygulama]
Global Boya(4) As String: Global BoyaTheme As String '[Görünüm]
Global Güven(5) As String '[Güvenlik]
Global NetPath(255) As String: Global NetName(255) As String
Global MyPath As String: Global MyName As String '[Að]
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub Main()
    'daha önceden açýlmýþsa bir daha açma
    If App.PrevInstance = True Then End
    'Gerekli dosyalar kontrol ediliyor.
    Call Dosya_kontrol
    'telif haklarý kontrolü yapýlýyor...
    Call Non_Changable_Telif_Check
    'deðiþkenler tanýmlanýr. settings.ini varsa okunur
    Call CheckLogging: Call VarsayýlanlaraDön: Call AyarlarýOku
    'logging
    Write #7, Day(Date) & "." & Month(Date) & "." & Year(Date), "md_Main", "Main", "Successful"
    'eðer açma parolsý girilecek ise girmesini iste aksi takdirde show MSN
    If Güven(0) = "1" Then frm_GetPass.Show Else frm_Messenger.Show
End Sub
Public Sub TheEnd()
    Write #7, Day(Date) & "." & Month(Date) & "." & Year(Date), "md_Main", "TheEnd", "Successful" 'logging
    Call AyarlarýKaydet: Unload frm_GetPass: Unload frm_Messenger: Close #7
    Kill App.Path & "\" & App.EXEName & ".exe.log"
    End
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+

Private Sub CheckLogging()
    Dim OpenClose, TimeOfSub, NameofForm, TypeOfSub, Description As String: OpenClose = AyarOku("Uygulama", "OpenClose", "Closed")
    If OpenClose = "Opened" And Dir(App.Path & "\" & App.EXEName & ".exe.log") <> "" Then
        Open App.Path & "\" & App.EXEName & ".exe.log" For Input As #7
        Open App.Path & "\crash." & App.EXEName & ".exe.log" For Append As #9
            Do While EOF(7) = False
                Input #7, TimeOfSub, NameofForm, TypeOfSub, Description
                Write #9, TimeOfSub, NameofForm, TypeOfSub, Description
            Loop
        Write #9, String(50, "-"): Close #9: Close #7
    End If
    AyarKaydet "Uygulama", "OpenClose", "Opened"
    Open App.Path & "\" & App.EXEName & ".exe.log" For Output As #7
End Sub
Public Sub VarsayýlanlaraDön()
    Dim i As Byte
    '[Konum]
    For i = 0 To 2: Konum(i) = "0": Next i: Kenar(0) = "5000": Kenar(1) = "8000"
    '[Uygulama]
    Ayar(0) = "0": Ayar(1) = "0": Ayar(2) = "0": Ayar(3) = App.Path & "\sounds\mesaj.mp3": Ayar(4) = "0": Ayar(5) = "0": Ayar(6) = "0"
    '[Görünüm]
    BoyaTheme = "Standart": For i = 0 To 4: Boya(i) = RenkTemalarý(BoyaTheme, i): Next i
    '[Güvenlik]
    Güven(0) = "0":  Güven(1) = "0": Güven(2) = "0": Güven(3) = "1": Güven(4) = CalculateMD5("1234"): Güven(5) = CalculateMD5("9876543210"):
    '[Að]
    MyPath = "c:\sohbet.txt": MyName = Environ("COMPUTERNAME")
End Sub
Private Sub AyarlarýOku()
    Dim i As Byte
    '[Konum]
    For i = 0 To 2: Konum(i) = AyarOku("Konum", "Konum(" & i & ")", CStr(Konum(i))): Next i
    For i = 0 To 1: Kenar(i) = AyarOku("Konum", "Kenar(" & i & ")", CStr(Kenar(i))): Next i
    '[Uygulama]
    For i = 0 To 6: Ayar(i) = AyarOku("Uygulama", "Ayar(" & i & ")", CStr(Ayar(i))): Next i
    '[Görünüm]
    For i = 0 To 4: Boya(i) = AyarOku("Görünüm", "Boya(" & i & ")", CStr(Boya(i))): Next i
    BoyaTheme = AyarOku("Görünüm", "BoyaTheme", CStr(BoyaTheme))
    '[Güvenlik]
    For i = 0 To 5: Güven(i) = AyarOku("Güvenlik", "Güven(" & i & ")", CStr(Güven(i))): Next i
    '[Að]
    Open App.Path & "\net.lst" For Input As #1: For i = 0 To 254
    If EOF(1) = True Then Exit For
    Input #1, NetName(i): NetPath(i) = "\\" & NetName(i) & "\c\sohbet.txt": Next i: Close #1
End Sub
Sub AyarlarýKaydet()
    Dim i As Byte
    '[Konum]
    For i = 0 To 2: AyarKaydet "Konum", "Konum(" & i & ")", CStr(Konum(i)): Next i
    For i = 0 To 1: AyarKaydet "Konum", "Kenar(" & i & ")", CStr(Kenar(i)): Next i
    AyarKaydet "Uygulama", "OpenClose", "Closed"
    '[Uygulama]
    For i = 0 To 6: AyarKaydet "Uygulama", "Ayar(" & i & ")", CStr(Ayar(i)): Next i
    '[Görünüm]
    For i = 0 To 4: AyarKaydet "Görünüm", "Boya(" & i & ")", CStr(Boya(i)): Next i
    AyarKaydet "Görünüm", "BoyaTheme", CStr(BoyaTheme)
    '[Güvenlik]
    For i = 0 To 5: AyarKaydet "Güvenlik", "Güven(" & i & ")", CStr(Güven(i)): Next i
    '[Að]
    Open App.Path & "\net.lst" For Output As #1: For i = 0 To 254
    If NetName(i) = "" Then Exit For
    Write #1, NetName(i): Next i: Close #1
End Sub

'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
Private Sub Non_Changable_Telif_Check()
    If App.CompanyName = "Sekmenler Tech." Then
        If App.LegalCopyright = "© " + App.CompanyName Then Exit Sub
    End If
    MsgBox "Uygulamanýn telif haklarý deðiþtirilmiþ." + Chr(13) + Chr(10) + "Lütfen uygulamayý tekrar kurun.", vbApplicationModal + vbCritical
    End
End Sub
Sub Dosya_kontrol()
    On Local Error Resume Next: Dim Dosyalarýn_Durumu_iyimi As Boolean: Dosyalarýn_Durumu_iyimi = True
    'sohbet.txt temizleniyor ve çevre deðiþkenleri yazýlýyor yoksa...
    Open "c:\sohbet.txt" For Output As #1: Close #1 'sohbet.txt
    If Dir(App.Path & "\net.lst") = "" Then Open App.Path & "\net.lst" For Output As #1: Close #1 'net.lst
    'iþte önemli dosya ve klasörler
    If Dir(App.Path + "\sounds\mesaj.mp3") = "" Then Dosyalarýn_Durumu_iyimi = False
    If Dir(App.Path + "\sounds\credits.avi") = "" Then Dosyalarýn_Durumu_iyimi = False
    'önemli dosyalar eksikse çalýþma!
    If Dosyalarýn_Durumu_iyimi = False Then MsgBox "Programýn dosyalarý hasar görmüþ. Lütfen programý tekrar yükleyin !", vbCritical + vbApplicationModal, "Hatalý Yükleme": End
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
