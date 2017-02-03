VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Kayitlar 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kayit Bilgisi"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   Icon            =   "Kayitlar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Kayit Ara"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6945
      TabIndex        =   7
      Top             =   2730
      Width           =   2010
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Kayitlar.frx":27A2
      Height          =   3975
      Left            =   540
      OleObjectBlob   =   "Kayitlar.frx":27B6
      TabIndex        =   13
      Top             =   4080
      Width           =   8415
   End
   Begin VB.TextBox Text4 
      Height          =   390
      Left            =   3473
      TabIndex        =   18
      Top             =   232
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Yeni"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   533
      TabIndex        =   4
      Top             =   2752
      Width           =   1400
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Son>>"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7763
      TabIndex        =   12
      Top             =   3487
      Width           =   1185
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ýleri>"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6518
      TabIndex        =   11
      Top             =   3487
      Width           =   1185
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<<Ýlk"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3788
      TabIndex        =   9
      Top             =   3487
      Width           =   1185
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<Geri"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5048
      TabIndex        =   10
      Top             =   3487
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kaydet"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2033
      TabIndex        =   3
      Top             =   2752
      Width           =   1400
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3473
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1807
      Width           =   5475
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3473
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1207
      Width           =   5475
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   3473
      MaxLength       =   50
      TabIndex        =   0
      Top             =   622
      Width           =   5475
   End
   Begin VB.Data Data1 
      BOFAction       =   1  'BOF
      Caption         =   "Veri Kaynaði"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   1  'EOF
      Exclusive       =   -1  'True
      Height          =   345
      Left            =   525
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "veriler"
      Top             =   2310
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sil"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5033
      TabIndex        =   6
      Top             =   2752
      Width           =   1400
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Düzenle"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3533
      TabIndex        =   5
      Top             =   2752
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anasayfa"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   533
      TabIndex        =   8
      Top             =   3487
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sicil Numarasi"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   18
         Charset         =   162
         Weight          =   800
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   390
      Left            =   608
      TabIndex        =   16
      Top             =   1807
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Doðum Tarihi"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   18
         Charset         =   162
         Weight          =   800
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   390
      Left            =   608
      TabIndex        =   15
      Top             =   1207
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Adi Soyadi"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   18
         Charset         =   162
         Weight          =   800
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   390
      Left            =   608
      TabIndex        =   14
      Top             =   622
      Width           =   1320
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Height          =   8100
      Left            =   225
      TabIndex        =   17
      Top             =   225
      Width           =   9255
   End
End
Attribute VB_Name = "Kayitlar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numara1, numara2 As Long
Dim sil, soru, kriter, ARA, cevap, m As String
Private Sub Command1_Click()
On Error Resume Next
Ýptal
Data1.Recordset.Close
Giriþ.Show: Unload Me
End Sub
Private Sub Command10_Click()
On Error GoTo hata
With Data1.Recordset
    .MoveFirst
    numara1 = 1: numara2 = 1
bas:
    If .EOF Then
        Text4 = numara1
    ElseIf .Fields("No") = numara1 Then
        .MoveNext: numara1 = numara1 + 1: numara2 = numara2 + 1: GoTo bas
    Else
        Text4 = numara1: GoTo son
    End If
son:
End With
dev:
Command1.Enabled = False: Command10.Enabled = False: Command3.Enabled = False: Command9.Enabled = False
Command2.Enabled = True
Text1.Enabled = True: Text2.Enabled = True: Text3.Enabled = True
Text1 = "": Text2 = "": Text3 = ""
Text1.SetFocus
Exit Sub
hata:
If Err.Number = 3021 Then Text4 = "1": GoTo dev Else MsgBox "Yeni Kayit Oluþturulamadi", vbExclamation, "Dosya Hatasi"
End Sub
Private Sub Command2_Click()
If Text1 = "" Then MsgBox "   Eksik Bilgi Girdiniz     ", vbExclamation, "Kayit Yapilamadi": Text1.SetFocus: GoTo son
If Text2 = "" Then MsgBox "   Eksik Bilgi Girdiniz     ", vbExclamation, "Kayit Yapilamadi": Text2.SetFocus: GoTo son
If Text3 = "" Then MsgBox "   Eksik Bilgi Girdiniz     ", vbExclamation, "Kayit Yapilamadi": Text3.SetFocus: GoTo son
Text1 = UCase(Text1)
If Command2.Caption = "Kabul Et" Then GoTo düzenle
If Command2.Caption = "Kaydet" Then GoTo kayit
kayit:
Data1.Recordset.AddNew
GoTo son
düzenle:
Data1.Recordset.Edit
son:
alanaeþitle: Data1.Refresh: Ýptal: Command2.Caption = "Kaydet": Command5_Click
End Sub
Private Sub Command3_Click()
On Error GoTo hata
If Text1 = "" And Text2 = "" And Text3 = "" Then Kayitsiz: GoTo son
sil = MsgBox("Bu Kiþiyi Silmek Ýstiyor Musunuz?", vbYesNo + vbExclamation, "Kiþi Silme")
If sil = 6 Then GoTo devam Else GoTo son
devam:
Data1.Recordset.Delete
Text1 = "": Text2 = "": Text3 = "": Text4 = "": Data1.Refresh
son:
Exit Sub
hata:
If Err.Number = 3021 Then
Kayitsiz
Else: MsgBox Err.Number
End If
End Sub
Private Sub Command4_Click()
On Error GoTo hata
Data1.Recordset.MovePrevious
ekranaeþitle
Exit Sub
hata:
If Err.Number = 3021 Then Kayitsiz
End Sub
Private Sub Command5_Click()
On Error GoTo hata
Data1.Recordset.MoveFirst
ekranaeþitle
Exit Sub
hata:
If Err.Number = 3021 Then Kayitsiz
End Sub
Private Sub Command6_Click()
On Error GoTo hata
Data1.Recordset.MoveNext
ekranaeþitle
Exit Sub
hata:
If Err.Number = 3021 Then Kayitsiz
End Sub
Private Sub command7_click()
On Error GoTo hata
Data1.Recordset.MoveLast
ekranaeþitle
Exit Sub
hata:
If Err.Number = 3021 Then Kayitsiz
End Sub
Private Sub Command8_Click()
On Error GoTo ErrorHandler
If Command8.Caption = "Kayit Ara" Then GoTo kayitbul
If Command8.Caption = "Reset" Then GoTo reset
kayitbul:
'arama hazirliklari
Text1 = "": Text2 = "": Text3 = "": Text4 = ""
m = Chr(34)
ARA = InputBox("BULAK ÝSTEDÝÐÝNÝZ KÝÞÝNÝN TAM ADINI YAZINIZ", "KAYIT ARAMA")
ARA = UCase(ARA)
kriter = "select * from veriler where [Ad Soyad] like" & m & ARA & m
'ilk aramanin sonuçlarini ekrana verir
Data1.RecordSource = kriter
Data1.Refresh: ekranaeþitle
If Text4 = "" Then MsgBox "  ARADIÐINIZ ÝSÝMLE BÝRÝ BULUNAMADI ": GoTo ErrorHandler
cevap = MsgBox("ARADIÐINIZ KAYIT BU MU?", 4, "KAYIT ARAMA")
    If cevap = 6 Then GoTo son Else GoTo reset
'aranan bu kayit deðilse döngüye sokar kayit bulunasiya kadar veya
'sona ulaþilincaya kadar tüm kayitlara bakar
'her bulduðu sonucu ekrana verir
    Do While Not Data1.Recordset.NoMatch
    Data1.Recordset.MoveNext
    ekranaeþitle
    cevap = MsgBox("ARADIÐINIZ KAYIT BU MU?", 4, "KAYIT ARAMA")
    If cevap = 6 Then GoTo son
'sona ulaþirsa bulunamadi diye mesaj verir
    If Data1.Recordset.EOF Then
        MsgBox "ARADIÐINIZ ÝSÝMLE BÝR YETKÝLÝ BULUNAMADI"
        GoTo ErrorHandler
    End If
    Loop
'Kayit bulnamadiðinda dbgrid normale döndürülür
ErrorHandler:
Data1.RecordSource = "veriler"
Data1.Refresh
Command5_Click
Exit Sub
'Son
son:
Command8.Caption = "Reset"
Exit Sub
'Herþeyi normale döndürür - eski haline yani!
reset:
Data1.RecordSource = "veriler"
Data1.Refresh
Command8.Caption = "Kayit Ara"
End Sub
Private Sub Command9_Click()
On Error GoTo hata
If Text1 = "" And Text2 = "" And Text3 = "" Then Kayitsiz: GoTo son
Text1.Enabled = True: Text2.Enabled = True: Text3.Enabled = True
Text1.SetFocus
Command1.Enabled = False
Command2.Enabled = True: Command10.Enabled = False: Command3.Enabled = False: Command9.Enabled = False
Command2.Caption = "Kabul Et"
son:
Exit Sub
hata:
If Err.Number = 3021 Then Kayitsiz
End Sub
Private Sub Form_Activate()
On Error GoTo hata
Data1.Recordset.MoveFirst
ekranaeþitle
Data1.Refresh
Exit Sub
hata:
If Err.Number = 3021 Then Kayitsiz
End Sub
Private Sub Form_Load()
Data1.DatabaseName = App.Path + "\3308.mdb"
End Sub
Private Sub Form_Unload(Cancel As Integer)
Command1_Click
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
If KeyAscii = 27 Then Ýptal
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus
If KeyAscii = 27 Then Ýptal
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command2_Click
If KeyAscii = 27 Then Ýptal
End Sub
