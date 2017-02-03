Attribute VB_Name = "Module"
Dim RasgeleSe�im As Byte
Dim �pucu, �imdikiSaat, �imdikiTarih, �imdikiAy, YeniAy As String
Sub Main()
'Description: Prevents than one instance of an application from running
If App.PrevInstance = True Then
End
Else
'form1i g�sterir ama dokundurtmaz
Form1.Show
'splash ekrana gelir ve 2 saniye sonra kaybolur
frmSplash.Show
frmSplash.Timer1.Enabled = True
End If
End Sub
Sub �pucuYaz()
'k���k bir kontrol: e�er ipucu zaman�nda de�i�mi� ise tekrar de�i�tirme!
If Form1.LabelKontrol.Caption = "1" Then Exit Sub
If Form2.LabelKontrol.Caption = "1" Then Exit Sub
If Form3.LabelKontrol.Caption = "1" Then Exit Sub
If Form4.LabelKontrol.Caption = "1" Then Exit Sub
If Form5.LabelKontrol.Caption = "1" Then Exit Sub

'ilk olarak rasgele bir say� belirlenir.
Randomize
RasgeleSe�im = Int((Rnd * 15) + 1)
'sonra o se�ilen say�ya denk gelen ipucu belirlenir.
Select Case RasgeleSe�im
Case 1
    ipucu = "Masa�st�nde iken F3 tu�una bas�ld���nda arama penceresi gelir..."
Case 2
    ipucu = "Tek Aya��n Havadayken Di�er Aya��n� da Havaya Kald�rd���nda Yere D��ersin."
Case 3
    ipucu = "Bir papa�an�n bir aya��n� �ekince ingilizce di�er aya��n� �ekince almanca konu�uyormu�. �ki aya�� �ekilince salak yere d��m��..."
Case 4
    ipucu = "Ayn� Anda Hem M�zik Dinleyip Hem de Film �zlemeye �al���rsan�z Sesleri Birbirine Kar���r."
Case 5
    ipucu = "Bir Hoparlore Sormu�lar; Senin Sesin Niye �ok ��k�yor? Hoparlor G�r�lt�den Duyamam��."
Case 6
    ipucu = "�ok Konu�unca Geveze Olursun."
Case 7
    ipucu = "Bir Vantilat�r So�uk Hava Alamay�nca So�uk Hava Vermezmi�"
Case 8
    ipucu = "Hacker Bilgisayar: Bilgisayar Sistemleri: Bili�im �r�nleri: Sat�� ve Teknik Servis"
Case 9
    ipucu = "On Milyon, Bir Araba ��in Ucuz Bir Mouse Pad ��in Pahal�d�r."
Case 10
    ipucu = "Microsoft Windows XP En G�venilir Ve En ��levsel ��letim Sistemidir"
Case 11
    ipucu = "Turbo+Power Bilgisayar� H�zl� Kapatmaya Yarar"
Case 12
    ipucu = "Bu Program �smail Sekmeno�lu ve H�seyin Sekmeno�lu Taraf�ndan Yaz�lm�� Olup Her Hakk� Kendilerinde Sakl�d�r"
Case 13
    ipucu = "Bu Program Muhasebe Bilgilerinizi En H�zl� �ekilde Girmeniz ��in Ayarlanm��t�r"
Case 14
    ipucu = "Klavyedeki Eg Tu�u Klavyenin �zel Tu�lar�n� Etkinle�tirir"
Case 15
    ipucu = "Bizi Siz Yaratt�n�z"
End Select

'�imdide se�ilen ipucunu formlardaki yerine yaz�lacak
Form1.lblipucu.Caption = ipucu
Form2.lblipucu.Caption = ipucu
Form3.lblipucu.Caption = ipucu
Form4.lblipucu.Caption = ipucu
Form5.lblipucu.Caption = ipucu
'ipucunun zaman�nda de�i�ti�ini belirtiyor.
Form1.LabelKontrol.Caption = "1"
Form2.LabelKontrol.Caption = "1"
Form3.LabelKontrol.Caption = "1"
Form4.LabelKontrol.Caption = "1"
Form5.LabelKontrol.Caption = "1"
End Sub
Sub MarkaYenile()
'comboyu temizler ve "<Bilinmeyen>" diye ekler hemen
Form1.Combo1.Clear
Form1.Combo1.AddItem "<Bilinmeyen>"
Open App.Path + "\data\cmd1.nfo" For Input As #1
'dosyadaki maddeler comboya eklenir
bas1:
If EOF(1) Then GoTo son1
Input #1, YeniMarka
Form1.Combo1.AddItem YeniMarka
GoTo bas1
son1:
'dosya kapat�l�p combodaki ilk ��e yani "<Bilinmeyen>" se�ilir
Close #1
Form1.Combo1.ListIndex = "0"
End Sub
Sub ModelYenile()
'comboyu temizler ve "<Bilinmeyen>" diye ekler hemen
Form1.Combo2.Clear
Form1.Combo2.AddItem "<Bilinmeyen>"
Open App.Path + "\data\cmd3.nfo" For Input As #1
'dosyadaki maddeler comboya eklenir
bas3:
If EOF(1) Then GoTo son3
Input #1, YeniModel
Form1.Combo2.AddItem YeniModel
GoTo bas3
son3:
'dosya kapat�l�p combodaki ilk ��e yani "<Bilinmeyen>" se�ilir
Close #1
Form1.Combo2.ListIndex = "0"
End Sub
Sub T�rYenile()
'comboyu temizler ve "<Bilinmeyen>" diye ekler hemen
Form1.Combo3.Clear
Form1.Combo3.AddItem "<Bilinmeyen>"
Open App.Path + "\data\cmd5.nfo" For Input As #1
'dosyadaki maddeler comboya eklenir
bas5:
If EOF(1) Then GoTo son5
Input #1, YeniT�r
Form1.Combo3.AddItem YeniT�r
GoTo bas5
son5:
'dosya kapat�l�p combodaki ilk ��e yani "<Bilinmeyen>" se�ilir
Close #1
Form1.Combo3.ListIndex = "0"
End Sub
Sub GarantiYenile()
'comboyu temizler ve "<Bilinmeyen>" diye ekler hemen
Form1.Combo4.Clear
Form1.Combo4.AddItem "<Bilinmeyen>"
Open App.Path + "\data\cmd7.nfo" For Input As #1
'dosyadaki maddeler comboya eklenir
bas7:
If EOF(1) Then GoTo son7
Input #1, YeniGaranti
Form1.Combo4.AddItem YeniGaranti
GoTo bas7
son7:
'dosya kapat�l�p combodaki ilk ��e yani "<Bilinmeyen>" se�ilir
Close #1
Form1.Combo4.ListIndex = "0"
End Sub
Sub ZamanBelirt()
'�u andaki zaman�n saatini ve dakikas�n� al�yoruz
�imdikiSaat = Left$(Time$, 5)
'i�inde bulundupumuz ay�n ka��nc� ay oldu�unu buluyoruz
�imdikiAy = Left$(Date$, 2)
'buldu�umuz ay�n ad�n� bri tan�ml� de�i�kene aktar�yoruz
Select Case �imdikiAy
Case 1
YeniAy = "Ocak"
Case 2
YeniAy = "�ubat"
Case 3
YeniAy = "Mart"
Case 4
YeniAy = "Nisan"
Case 5
YeniAy = "May�s"
Case 6
YeniAy = "Haziran"
Case 7
YeniAy = "Temmuz"
Case 8
YeniAy = "A�ustos"
Case 9
YeniAy = "Eyl�l"
Case 10
YeniAy = "Ekim"
Case 11
YeniAy = "Kas�m"
Case 12
YeniAy = "Aral�k"
End Select
'tarihi do�ru d�zg�n belirtiyoruz
�imdikiTarih = Mid$(Date$, 4, 2) + " " + YeniAy + " " + Right$(Date$, 4)
'�imdide tarih ve saat istenilen yerlere yazd�r�yoruz
Form1.TxtSaat = �imdikiSaat: Form1.TxtTarih = �imdikiTarih
Form2.Text3 = �imdikiSaat: Form2.Text2 = �imdikiTarih
End Sub
Sub Malzemeler()
Open App.Path + "\stuff\market.mlz" For Input As #1
Bas:
If EOF(1) Then GoTo Son
'mrk=marka, mdl=model, cns=cins, grn=garanti, ftr=faturatarihi,fno=fatura no,
'�zl=�zellikleri, bsn=barkod*seri no, fyt=fiyat�, atr=al�m tarihi, ast=al�m saati
Input #1, mrk, mdl, cns, grn, ftr, fno, �zl, bsn, fyt, atr, ast
'sat��listesi -sat�labilecek mallar
Form2.List1.AddItem cns & "   " & mdl & " " & mrk & "  Garantisi :  " & grn
GoTo Bas
Son:
Close #1
End Sub
Sub Ayr�nt�Getir()
Form2.ListAyr�nt�.Clear
Open App.Path + "\stuff\market.mlz" For Input As #1
Bas:
If EOF(1) Then GoTo Son
'mrk=marka, mdl=model, cns=cins, grn=garanti, ftr=faturatarihi,fno=fatura no,
'�zl=�zellikleri, bsn=barkod*seri no, fyt=fiyat�, atr=al�m tarihi, ast=al�m saati
Input #1, mrk, mdl, cns, grn, ftr, fno, �zl, fyt, bsn, atr, ast
'sat��listesi -sat�labilecek mallar
If Form2.List1.Text = cns & "   " & mdl & " " & mrk & "  Garantisi :  " & grn Then
Form2.ListAyr�nt�.AddItem "�zellikleri:     " & �zl
Form2.ListAyr�nt�.AddItem "  Fiyat�:   " & fyt & "    " & "  Seri No :  " & bsn
Form2.ListAyr�nt�.AddItem "    Fatura Tarihi :  " & ftr & "    Fatura no :  " & fno
Form2.ListAyr�nt�.AddItem "      Al�m Tarihi :  " & atr & "    Al�m Saati :  " & ast
Else
GoTo Bas
End If
Son:
Close #1
End Sub
