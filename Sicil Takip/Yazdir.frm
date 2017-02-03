VERSION 5.00
Begin VB.Form Yazdir 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yazici Çiktisi"
   ClientHeight    =   2370
   ClientLeft      =   -4305
   ClientTop       =   -2115
   ClientWidth     =   7365
   Icon            =   "Yazdir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   420
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Yazici"
      Top             =   420
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   420
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "veriler"
      Top             =   735
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Prim Listesi"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   3255
      TabIndex        =   3
      Top             =   1250
      Width           =   3585
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sigortaya Giden Liste"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   3255
      TabIndex        =   2
      Top             =   400
      Width           =   3585
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
      Left            =   400
      TabIndex        =   1
      Top             =   1465
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Height          =   1900
      Left            =   195
      TabIndex        =   0
      Top             =   195
      Width           =   6900
   End
End
Attribute VB_Name = "Yazdir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gün, Ücret, yaþi, sira As Long
Dim yil, dogay, tarihay As String
Private Sub Command1_Click()
On Error Resume Next
Data1.Recordset.Close: Data2.Recordset.Close
Giriþ.Show: Unload Me
End Sub
Private Sub Command2_Click()
'On Error GoTo hata
gün = InputBox("Çaliþilan Gün Sayisini Girin")
'ücretleri öðrenir
Open App.Path + "\salary.sys" For Input As #1
If EOF(1) Then MsgBox "Ücretler Belirtilmemiþ": Exit Sub
Input #1, a, b, c, d
alti = a: üstü = b
Close #1
'ücretler öðrenildi
'burada data.recordseti temizler
ilk:
If Data2.Recordset.EOF = True Then GoTo iki
Data2.Recordset.MoveFirst
Data2.Recordset.Delete
Data2.Recordset.MoveNext
GoTo ilk
iki:
'sirada sigorta listesi
'her öðrenci için teker teker liste girdisi hazirlar
Data1.Recordset.MoveFirst
sira = 0
bas:
Data2.Recordset.AddNew
If Data1.Recordset.EOF Then GoTo son
If Data1.Recordset.Fields("No") = "" Then Data1.Recordset.MoveNext
'Data2.Recordset.Fields("No") = Data1.Recordset.Fields("No")
Data2.Recordset.Fields("Ad Soyad") = Data1.Recordset.Fields("Ad Soyad")
Data2.Recordset.Fields("Sicil No") = Data1.Recordset.Fields("Sicil No")
'yaþ hesaplama bölümü
yil = Val(Right(Data1.Recordset.Fields("Doðum Tarihi"), 4))
dogay = Val(Mid(Data1.Recordset.Fields("Doðum Tarihi"), 4, 2))
tarihay = Val(Left(Date$, 2))
If tarihay > dogay Then
yaþi = Val(Right(Date$, 4)) - yil
Else
yaþi = Val(Right(Date$, 4) - 1) - yil
End If
'yaþina göre ücretini hesaplar
If yaþi >= 16 Then Ücret = üstü Else Ücret = alti
'listeye girdi hazirlar
Data2.Recordset.Fields("Toplam") = Ücret * gün
Data2.Recordset.Fields("GTop") = Data2.Recordset.Fields("GTop") + Data2.Recordset.Fields("Toplam")
Data2.Recordset.Fields("Gün") = gün
sira = sira + 1
Data1.Recordset.MoveNext
GoTo bas
son:
Data2.Recordset.Fields("GunTop") = sira * gün
DataReport1.Show
Exit Sub
hata:
End Sub
Private Sub Form_Load()
Data1.DatabaseName = App.Path + "\3308.mdb"
Data2.DatabaseName = App.Path + "\3308.mdb"
End Sub
Private Sub Form_Unload(Cancel As Integer)
Command1_Click
End Sub
