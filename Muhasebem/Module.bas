Attribute VB_Name = "Module"
Dim RasgeleSeçim As Byte
Dim Ýpucu, ÞimdikiSaat, ÞimdikiTarih, ÞimdikiAy, YeniAy As String
Sub Main()
'Description: Prevents than one instance of an application from running
If App.PrevInstance = True Then
End
Else
'form1i gösterir ama dokundurtmaz
Form1.Show
'splash ekrana gelir ve 2 saniye sonra kaybolur
frmSplash.Show
frmSplash.Timer1.Enabled = True
End If
End Sub
Sub ÝpucuYaz()
'küçük bir kontrol: eðer ipucu zamanýnda deðiþmiþ ise tekrar deðiþtirme!
If Form1.LabelKontrol.Caption = "1" Then Exit Sub
If Form2.LabelKontrol.Caption = "1" Then Exit Sub
If Form3.LabelKontrol.Caption = "1" Then Exit Sub
If Form4.LabelKontrol.Caption = "1" Then Exit Sub
If Form5.LabelKontrol.Caption = "1" Then Exit Sub

'ilk olarak rasgele bir sayý belirlenir.
Randomize
RasgeleSeçim = Int((Rnd * 15) + 1)
'sonra o seçilen sayýya denk gelen ipucu belirlenir.
Select Case RasgeleSeçim
Case 1
    ipucu = "Masaüstünde iken F3 tuþuna basýldýðýnda arama penceresi gelir..."
Case 2
    ipucu = "Tek Ayaðýn Havadayken Diðer Ayaðýný da Havaya Kaldýrdýðýnda Yere Düþersin."
Case 3
    ipucu = "Bir papaðanýn bir ayaðýný çekince ingilizce diðer ayaðýný çekince almanca konuþuyormuþ. Ýki ayaðý çekilince salak yere düþmüþ..."
Case 4
    ipucu = "Ayný Anda Hem Müzik Dinleyip Hem de Film Ýzlemeye Çalýþýrsanýz Sesleri Birbirine Karýþýr."
Case 5
    ipucu = "Bir Hoparlore Sormuþlar; Senin Sesin Niye Çok Çýkýyor? Hoparlor Gürültüden Duyamamýþ."
Case 6
    ipucu = "Çok Konuþunca Geveze Olursun."
Case 7
    ipucu = "Bir Vantilatör Soðuk Hava Alamayýnca Soðuk Hava Vermezmiþ"
Case 8
    ipucu = "Hacker Bilgisayar: Bilgisayar Sistemleri: Biliþim Ürünleri: Satýþ ve Teknik Servis"
Case 9
    ipucu = "On Milyon, Bir Araba Ýçin Ucuz Bir Mouse Pad Ýçin Pahalýdýr."
Case 10
    ipucu = "Microsoft Windows XP En Güvenilir Ve En Ýþlevsel Ýþletim Sistemidir"
Case 11
    ipucu = "Turbo+Power Bilgisayarý Hýzlý Kapatmaya Yarar"
Case 12
    ipucu = "Bu Program Ýsmail Sekmenoðlu ve Hüseyin Sekmenoðlu Tarafýndan Yazýlmýþ Olup Her Hakký Kendilerinde Saklýdýr"
Case 13
    ipucu = "Bu Program Muhasebe Bilgilerinizi En Hýzlý Þekilde Girmeniz Ýçin Ayarlanmýþtýr"
Case 14
    ipucu = "Klavyedeki Eg Tuþu Klavyenin Özel Tuþlarýný Etkinleþtirir"
Case 15
    ipucu = "Bizi Siz Yarattýnýz"
End Select

'þimdide seçilen ipucunu formlardaki yerine yazýlacak
Form1.lblipucu.Caption = ipucu
Form2.lblipucu.Caption = ipucu
Form3.lblipucu.Caption = ipucu
Form4.lblipucu.Caption = ipucu
Form5.lblipucu.Caption = ipucu
'ipucunun zamanýnda deðiþtiðini belirtiyor.
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
'dosya kapatýlýp combodaki ilk öðe yani "<Bilinmeyen>" seçilir
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
'dosya kapatýlýp combodaki ilk öðe yani "<Bilinmeyen>" seçilir
Close #1
Form1.Combo2.ListIndex = "0"
End Sub
Sub TürYenile()
'comboyu temizler ve "<Bilinmeyen>" diye ekler hemen
Form1.Combo3.Clear
Form1.Combo3.AddItem "<Bilinmeyen>"
Open App.Path + "\data\cmd5.nfo" For Input As #1
'dosyadaki maddeler comboya eklenir
bas5:
If EOF(1) Then GoTo son5
Input #1, YeniTür
Form1.Combo3.AddItem YeniTür
GoTo bas5
son5:
'dosya kapatýlýp combodaki ilk öðe yani "<Bilinmeyen>" seçilir
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
'dosya kapatýlýp combodaki ilk öðe yani "<Bilinmeyen>" seçilir
Close #1
Form1.Combo4.ListIndex = "0"
End Sub
Sub ZamanBelirt()
'þu andaki zamanýn saatini ve dakikasýný alýyoruz
ÞimdikiSaat = Left$(Time$, 5)
'içinde bulundupumuz ayýn kaçýncý ay olduðunu buluyoruz
ÞimdikiAy = Left$(Date$, 2)
'bulduðumuz ayýn adýný bri tanýmlý deðiþkene aktarýyoruz
Select Case ÞimdikiAy
Case 1
YeniAy = "Ocak"
Case 2
YeniAy = "Þubat"
Case 3
YeniAy = "Mart"
Case 4
YeniAy = "Nisan"
Case 5
YeniAy = "Mayýs"
Case 6
YeniAy = "Haziran"
Case 7
YeniAy = "Temmuz"
Case 8
YeniAy = "Aðustos"
Case 9
YeniAy = "Eylül"
Case 10
YeniAy = "Ekim"
Case 11
YeniAy = "Kasým"
Case 12
YeniAy = "Aralýk"
End Select
'tarihi doðru düzgün belirtiyoruz
ÞimdikiTarih = Mid$(Date$, 4, 2) + " " + YeniAy + " " + Right$(Date$, 4)
'þimdide tarih ve saat istenilen yerlere yazdýrýyoruz
Form1.TxtSaat = ÞimdikiSaat: Form1.TxtTarih = ÞimdikiTarih
Form2.Text3 = ÞimdikiSaat: Form2.Text2 = ÞimdikiTarih
End Sub
Sub Malzemeler()
Open App.Path + "\stuff\market.mlz" For Input As #1
Bas:
If EOF(1) Then GoTo Son
'mrk=marka, mdl=model, cns=cins, grn=garanti, ftr=faturatarihi,fno=fatura no,
'özl=özellikleri, bsn=barkod*seri no, fyt=fiyatý, atr=alým tarihi, ast=alým saati
Input #1, mrk, mdl, cns, grn, ftr, fno, özl, bsn, fyt, atr, ast
'satýþlistesi -satýlabilecek mallar
Form2.List1.AddItem cns & "   " & mdl & " " & mrk & "  Garantisi :  " & grn
GoTo Bas
Son:
Close #1
End Sub
Sub AyrýntýGetir()
Form2.ListAyrýntý.Clear
Open App.Path + "\stuff\market.mlz" For Input As #1
Bas:
If EOF(1) Then GoTo Son
'mrk=marka, mdl=model, cns=cins, grn=garanti, ftr=faturatarihi,fno=fatura no,
'özl=özellikleri, bsn=barkod*seri no, fyt=fiyatý, atr=alým tarihi, ast=alým saati
Input #1, mrk, mdl, cns, grn, ftr, fno, özl, fyt, bsn, atr, ast
'satýþlistesi -satýlabilecek mallar
If Form2.List1.Text = cns & "   " & mdl & " " & mrk & "  Garantisi :  " & grn Then
Form2.ListAyrýntý.AddItem "Özellikleri:     " & özl
Form2.ListAyrýntý.AddItem "  Fiyatý:   " & fyt & "    " & "  Seri No :  " & bsn
Form2.ListAyrýntý.AddItem "    Fatura Tarihi :  " & ftr & "    Fatura no :  " & fno
Form2.ListAyrýntý.AddItem "      Alým Tarihi :  " & atr & "    Alým Saati :  " & ast
Else
GoTo Bas
End If
Son:
Close #1
End Sub
