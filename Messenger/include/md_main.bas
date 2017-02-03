Attribute VB_Name = "md_Main"
Option Explicit
'de�i�kenler
Global Konum(2) As String: Global Kenar(1) As String '[Konum]
Global Ayar(6) As String '[Uygulama]
Global Boya(4) As String: Global BoyaTheme As String '[G�r�n�m]
Global G�ven(5) As String '[G�venlik]
Global NetPath(255) As String: Global NetName(255) As String
Global MyPath As String: Global MyName As String '[A�]
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub Main()
    'daha �nceden a��lm��sa bir daha a�ma
    If App.PrevInstance = True Then End
    'Gerekli dosyalar kontrol ediliyor.
    Call Dosya_kontrol
    'telif haklar� kontrol� yap�l�yor...
    Call Non_Changable_Telif_Check
    'de�i�kenler tan�mlan�r. settings.ini varsa okunur
    Call CheckLogging: Call Varsay�lanlaraD�n: Call Ayarlar�Oku
    'logging
    Write #7, Day(Date) & "." & Month(Date) & "." & Year(Date), "md_Main", "Main", "Successful"
    'e�er a�ma parols� girilecek ise girmesini iste aksi takdirde show MSN
    If G�ven(0) = "1" Then frm_GetPass.Show Else frm_Messenger.Show
End Sub
Public Sub TheEnd()
    Write #7, Day(Date) & "." & Month(Date) & "." & Year(Date), "md_Main", "TheEnd", "Successful" 'logging
    Call Ayarlar�Kaydet: Unload frm_GetPass: Unload frm_Messenger: Close #7
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
Public Sub Varsay�lanlaraD�n()
    Dim i As Byte
    '[Konum]
    For i = 0 To 2: Konum(i) = "0": Next i: Kenar(0) = "5000": Kenar(1) = "8000"
    '[Uygulama]
    Ayar(0) = "0": Ayar(1) = "0": Ayar(2) = "0": Ayar(3) = App.Path & "\sounds\mesaj.mp3": Ayar(4) = "0": Ayar(5) = "0": Ayar(6) = "0"
    '[G�r�n�m]
    BoyaTheme = "Standart": For i = 0 To 4: Boya(i) = RenkTemalar�(BoyaTheme, i): Next i
    '[G�venlik]
    G�ven(0) = "0":  G�ven(1) = "0": G�ven(2) = "0": G�ven(3) = "1": G�ven(4) = CalculateMD5("1234"): G�ven(5) = CalculateMD5("9876543210"):
    '[A�]
    MyPath = "c:\sohbet.txt": MyName = Environ("COMPUTERNAME")
End Sub
Private Sub Ayarlar�Oku()
    Dim i As Byte
    '[Konum]
    For i = 0 To 2: Konum(i) = AyarOku("Konum", "Konum(" & i & ")", CStr(Konum(i))): Next i
    For i = 0 To 1: Kenar(i) = AyarOku("Konum", "Kenar(" & i & ")", CStr(Kenar(i))): Next i
    '[Uygulama]
    For i = 0 To 6: Ayar(i) = AyarOku("Uygulama", "Ayar(" & i & ")", CStr(Ayar(i))): Next i
    '[G�r�n�m]
    For i = 0 To 4: Boya(i) = AyarOku("G�r�n�m", "Boya(" & i & ")", CStr(Boya(i))): Next i
    BoyaTheme = AyarOku("G�r�n�m", "BoyaTheme", CStr(BoyaTheme))
    '[G�venlik]
    For i = 0 To 5: G�ven(i) = AyarOku("G�venlik", "G�ven(" & i & ")", CStr(G�ven(i))): Next i
    '[A�]
    Open App.Path & "\net.lst" For Input As #1: For i = 0 To 254
    If EOF(1) = True Then Exit For
    Input #1, NetName(i): NetPath(i) = "\\" & NetName(i) & "\c\sohbet.txt": Next i: Close #1
End Sub
Sub Ayarlar�Kaydet()
    Dim i As Byte
    '[Konum]
    For i = 0 To 2: AyarKaydet "Konum", "Konum(" & i & ")", CStr(Konum(i)): Next i
    For i = 0 To 1: AyarKaydet "Konum", "Kenar(" & i & ")", CStr(Kenar(i)): Next i
    AyarKaydet "Uygulama", "OpenClose", "Closed"
    '[Uygulama]
    For i = 0 To 6: AyarKaydet "Uygulama", "Ayar(" & i & ")", CStr(Ayar(i)): Next i
    '[G�r�n�m]
    For i = 0 To 4: AyarKaydet "G�r�n�m", "Boya(" & i & ")", CStr(Boya(i)): Next i
    AyarKaydet "G�r�n�m", "BoyaTheme", CStr(BoyaTheme)
    '[G�venlik]
    For i = 0 To 5: AyarKaydet "G�venlik", "G�ven(" & i & ")", CStr(G�ven(i)): Next i
    '[A�]
    Open App.Path & "\net.lst" For Output As #1: For i = 0 To 254
    If NetName(i) = "" Then Exit For
    Write #1, NetName(i): Next i: Close #1
End Sub

'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
Private Sub Non_Changable_Telif_Check()
    If App.CompanyName = "Sekmenler Tech." Then
        If App.LegalCopyright = "� " + App.CompanyName Then Exit Sub
    End If
    MsgBox "Uygulaman�n telif haklar� de�i�tirilmi�." + Chr(13) + Chr(10) + "L�tfen uygulamay� tekrar kurun.", vbApplicationModal + vbCritical
    End
End Sub
Sub Dosya_kontrol()
    On Local Error Resume Next: Dim Dosyalar�n_Durumu_iyimi As Boolean: Dosyalar�n_Durumu_iyimi = True
    'sohbet.txt temizleniyor ve �evre de�i�kenleri yaz�l�yor yoksa...
    Open "c:\sohbet.txt" For Output As #1: Close #1 'sohbet.txt
    If Dir(App.Path & "\net.lst") = "" Then Open App.Path & "\net.lst" For Output As #1: Close #1 'net.lst
    'i�te �nemli dosya ve klas�rler
    If Dir(App.Path + "\sounds\mesaj.mp3") = "" Then Dosyalar�n_Durumu_iyimi = False
    If Dir(App.Path + "\sounds\credits.avi") = "" Then Dosyalar�n_Durumu_iyimi = False
    '�nemli dosyalar eksikse �al��ma!
    If Dosyalar�n_Durumu_iyimi = False Then MsgBox "Program�n dosyalar� hasar g�rm��. L�tfen program� tekrar y�kleyin !", vbCritical + vbApplicationModal, "Hatal� Y�kleme": End
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
