VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form Siparis 
   BackColor       =   &H0096E06D&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7965
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSiparis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   7965
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3870
      Top             =   9285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0096E06D&
      Caption         =   "Resim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   134
      TabIndex        =   40
      Top             =   6915
      Width           =   7575
      Begin OsenXPCntrl.OsenXPButton olustur 
         Height          =   390
         Left            =   3885
         TabIndex        =   18
         Top             =   300
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
         BTYPE           =   3
         TX              =   "Olustur"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12648384
         BCOLO           =   12648384
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmSiparis.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox resim 
         BackColor       =   &H0096E06D&
         Caption         =   "Resim Kullan"
         Height          =   390
         Left            =   210
         TabIndex        =   16
         Top             =   300
         Width           =   1755
      End
      Begin OsenXPCntrl.OsenXPButton g�zat 
         Height          =   390
         Left            =   2100
         TabIndex        =   17
         Top             =   300
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
         BTYPE           =   3
         TX              =   "G�zat"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12648384
         BCOLO           =   12648384
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmSiparis.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton g�sterresim 
         Height          =   390
         Left            =   5670
         TabIndex        =   19
         Top             =   300
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   688
         BTYPE           =   3
         TX              =   "G�ster"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12648384
         BCOLO           =   12648384
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmSiparis.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label ResAdres 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   390
         Left            =   2100
         TabIndex        =   42
         Top             =   300
         Visible         =   0   'False
         Width           =   5295
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4350
      Top             =   9285
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0096E06D&
      Caption         =   "M�steri Bilgileri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   134
      TabIndex        =   35
      Top             =   1335
      Width           =   7575
      Begin VB.ComboBox soyad� 
         Height          =   315
         Left            =   2625
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   2200
      End
      Begin VB.ComboBox ad� 
         Height          =   315
         Left            =   210
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   630
         Width           =   2200
      End
      Begin VB.TextBox tel 
         Height          =   285
         Left            =   5040
         TabIndex        =   2
         Top             =   630
         Width           =   2200
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefonu"
         Height          =   195
         Left            =   5040
         TabIndex        =   38
         Top             =   315
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Soyad�"
         Height          =   195
         Left            =   2625
         TabIndex        =   37
         Top             =   315
         Width           =   480
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ad�"
         Height          =   195
         Left            =   225
         TabIndex        =   36
         Top             =   315
         Width           =   225
      End
   End
   Begin OsenXPCntrl.OsenXPButton Command1 
      Cancel          =   -1  'True
      Height          =   840
      Left            =   127
      TabIndex        =   21
      Top             =   7995
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1482
      BTYPE           =   3
      TX              =   "Geri"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmSiparis.frx":0060
      PICN            =   "frmSiparis.frx":007C
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton Command2 
      Default         =   -1  'True
      Height          =   840
      Left            =   5302
      TabIndex        =   20
      Top             =   7995
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1482
      BTYPE           =   3
      TX              =   "Kaydet"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmSiparis.frx":04CE
      PICN            =   "frmSiparis.frx":04EA
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0096E06D&
      Caption         =   "�s Bilgileri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4305
      Left            =   134
      TabIndex        =   22
      Top             =   2535
      Width           =   7575
      Begin VB.TextBox fiyat 
         Height          =   285
         Left            =   5115
         TabIndex        =   14
         Top             =   1305
         Width           =   2200
      End
      Begin VB.ComboBox cins 
         Height          =   315
         ItemData        =   "frmSiparis.frx":093C
         Left            =   5115
         List            =   "frmSiparis.frx":093E
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   825
         Width           =   2200
      End
      Begin VB.ComboBox kumas 
         Height          =   315
         ItemData        =   "frmSiparis.frx":0940
         Left            =   5115
         List            =   "frmSiparis.frx":0942
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   2200
      End
      Begin VB.TextBox acik 
         Height          =   1890
         Left            =   3885
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   2250
         Width           =   3465
      End
      Begin VB.TextBox bel 
         Height          =   285
         Left            =   1395
         TabIndex        =   6
         Top             =   2055
         Width           =   2200
      End
      Begin VB.TextBox basen 
         Height          =   285
         Left            =   1395
         TabIndex        =   7
         Top             =   2415
         Width           =   2200
      End
      Begin VB.TextBox gogus 
         Height          =   285
         Left            =   1395
         TabIndex        =   8
         Top             =   2775
         Width           =   2200
      End
      Begin VB.TextBox omuz 
         Height          =   285
         Left            =   1395
         TabIndex        =   9
         Top             =   3135
         Width           =   2200
      End
      Begin VB.TextBox kol 
         Height          =   285
         Left            =   1395
         TabIndex        =   10
         Top             =   3495
         Width           =   2200
      End
      Begin VB.TextBox boy 
         Height          =   285
         Left            =   1395
         TabIndex        =   11
         Top             =   3855
         Width           =   2200
      End
      Begin MSComCtl2.DTPicker sip 
         Height          =   315
         Left            =   1395
         TabIndex        =   3
         Top             =   360
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20840449
         CurrentDate     =   38459
      End
      Begin MSComCtl2.DTPicker pro 
         Height          =   315
         Left            =   1395
         TabIndex        =   4
         Top             =   825
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20840449
         CurrentDate     =   38459
      End
      Begin MSComCtl2.DTPicker tes 
         Height          =   315
         Left            =   1395
         TabIndex        =   5
         Top             =   1290
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20840449
         CurrentDate     =   38459
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fiyat� (YTL)"
         Height          =   195
         Left            =   3885
         TabIndex        =   45
         Top             =   1350
         Width           =   795
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A��klama   :"
         Height          =   195
         Left            =   3885
         TabIndex        =   34
         Top             =   1980
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Siparis Tarihi"
         Height          =   195
         Left            =   225
         TabIndex        =   33
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prova Tarihi"
         Height          =   195
         Left            =   225
         TabIndex        =   32
         Top             =   885
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teslim Tarihi"
         Height          =   195
         Left            =   225
         TabIndex        =   31
         Top             =   1350
         Width           =   885
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kumas T�r�"
         Height          =   195
         Left            =   3885
         TabIndex        =   30
         Top             =   420
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   150
         X2              =   7245
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Elbise Cinsi"
         Height          =   195
         Left            =   3885
         TabIndex        =   29
         Top             =   885
         Width           =   795
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   3765
         X2              =   3765
         Y1              =   315
         Y2              =   4095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bel"
         Height          =   195
         Left            =   225
         TabIndex        =   28
         Top             =   2100
         Width           =   225
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basen"
         Height          =   195
         Left            =   225
         TabIndex        =   27
         Top             =   2460
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G���s"
         Height          =   195
         Left            =   225
         TabIndex        =   26
         Top             =   2820
         Width           =   465
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Omuz"
         Height          =   195
         Left            =   225
         TabIndex        =   25
         Top             =   3180
         Width           =   405
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kol"
         Height          =   195
         Left            =   225
         TabIndex        =   24
         Top             =   3540
         Width           =   225
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Boy"
         Height          =   195
         Left            =   225
         TabIndex        =   23
         Top             =   3900
         Width           =   270
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dikim Evi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3712
      TabIndex        =   44
      Top             =   720
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1110
      Left            =   217
      Picture         =   "frmSiparis.frx":0944
      Top             =   120
      Width           =   3555
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "S�PAR�� S�H�RBAZI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5062
      TabIndex        =   43
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label kayno2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�s Say�s� : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2647
      TabIndex        =   41
      Top             =   8475
      Width           =   2625
   End
   Begin VB.Label kayno 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M�steri Say�s� : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2647
      TabIndex        =   39
      Top             =   7995
      Width           =   2625
   End
End
Attribute VB_Name = "Siparis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ResimSayisi As String: Dim GeriAl As Boolean
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub Command2_Click()
    Write #7, Time$, Me.Name, "Command2_Click", "Start" 'logging
    On Local Error Resume Next
    Dim Control As Control: Dim tmp_Say�, BosYerVarM� As Boolean: Dim i As Integer
    Dim M�steriID, SipT�rID, KumasT�rID As String
    BosYerVarM� = False
    If Trim(ad�.Text) = "" Or Trim(soyad�.Text) = "" Or Trim(fiyat.Text) = "" Or Val(fiyat.Text) = 0 Then MsgBox "L�tfen m�sterinin ad�n�, soyad�n� ve mal�n fiyat�n� belirtiniz.", vbExclamation: Exit Sub
    'bosluklar temizleniyor
    For Each Control In Me
        If TypeOf Control Is TextBox Then
            Control.Text = Trim(Control.Text)
            If Control.Text = "" Then BosYerVarM� = True
        End If
    Next Control
    'bos yer varsa devam edilip edilmeyece�i soruluyor.
    If BosYerVarM� = True Then
        If MsgBox("Eksik Bilgi ��eriyor Devam Edilsin mi?", 36) = vbNo Then Exit Sub
    End If
    Me.Enabled = False: Me.MousePointer = 11
    'e�er resim eklenmisse, o resim pictures klas�r�ne kopyalan�yor.
    If ResAdres.Caption = "" Then resim.Value = 0
    If resim.Value = 1 Then
ilkResim:
        ResimNO = ResimNO + 1
        If Dir(App.path + "\Pictures\Pictures" & CStr(ResimNO) & ".jpg") <> "" Then GoTo ilkResim
        FileCopy ResAdres.Caption, App.path + "\Pictures\Pictures" & CStr(ResimNO) & ".jpg"
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''KAYIT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'gerekli �n incelemeler bitti�ine g�re kayda baslayabiliriz.
    'ilk �nce bu ada ve soyada sahiip bir kaydolmusmu bakaca��z.
    M�steriID = VarsaM�steriIDBul(ad�.Text, soyad�.Text) 'kaydolmus ise bize bir m�steri kodu verecek
    With Anasayfa.dtM�steri.Recordset
        If M�steriID = "0" Then 'daha �nce kaydolmam�s...
            .AddNew
            .Fields("Musteri_Adi") = ad�.Text
            .Fields("Musteri_Soyadi") = soyad�.Text
            .Fields("Musteri_Telefon") = tel.Text
            .Update
            M�steriID = VarsaM�steriIDBul(ad�.Text, soyad�.Text) 'kaydoldu�una g�re m�steriID'sini bulur.
        Else
            .Edit
            .Fields("Musteri_Telefon") = tel.Text
            .Update
        End If
    End With
    'siparis T�rleri
    tmp_Say� = False
    With Anasayfa.dtSiparisT�r�.Recordset
        If .RecordCount <> 0 Then
            .MoveFirst
            For i = 1 To .RecordCount  'burada bu siparis t�r�n�n daha �nce kaydolup olmad���n� arast�r�yoruz.
                If .Fields("Siparis_Adi") = cins.Text Then tmp_Say� = True: Exit For  'kaydolmussa tmp_say�=1 oluyor!
                .MoveNext
            Next i
        End If
        If tmp_Say� = False Then
            .AddNew
            .Fields("Siparis_Adi") = cins.Text
            .Update
            .MoveLast
            SipT�rID = Val(.Fields("Siparis_Turleri"))
        Else
            SipT�rID = Val(.Fields("Siparis_Turleri"))
        End If
    End With
    'kumas T�rleri
    tmp_Say� = False
    With Anasayfa.dtKumasT�r�.Recordset
        If .RecordCount <> 0 Then
            .MoveFirst
            For i = 1 To .RecordCount  'burada bu kumas t�r�n�n daha �nce kaydolup olmad���n� arast�r�yoruz.
                If .Fields("Kumas_Adi") = kumas.Text Then tmp_Say� = True: Exit For  'kaydolmussa tmp_say�=1 oluyor!
                .MoveNext
            Next i
        End If
        If tmp_Say� = False Then
            .AddNew
            .Fields("Kumas_Adi") = kumas.Text
            .Update
            .MoveLast
            KumasT�rID = Val(.Fields("Kumas_Turu"))
        Else
            KumasT�rID = Val(.Fields("Kumas_Turu"))
        End If
    End With
    's�ra geldi as�l ise, yani siparis datas�na...
    With Anasayfa.dtSiparis.Recordset
        .AddNew
        .Fields("Musteri_Kodu") = M�steriID
        .Fields("Siparis_Turu") = SipT�rID
        .Fields("Kumas_Turu") = KumasT�rID
        .Fields("Siparis_Tarihi") = sip.Value
        .Fields("Prova_Tarihi") = pro.Value
        .Fields("Teslim_Tarihi") = tes.Value
        .Fields("Aciklama") = acik.Text
        .Fields("Ucret") = fiyat.Text
        .Fields("Durum") = "Bitmedi"
        .Fields("Boy") = boy.Text
        .Fields("Basen") = basen.Text
        .Fields("Bel") = bel.Text
        .Fields("G���s") = gogus.Text
        .Fields("Omuz") = omuz.Text
        .Fields("Kol") = kol.Text
        .Update
        .MoveLast
        i = Val(.Fields("Siparis_Kodu"))
    End With
    If resim.Value = 1 Then
        With Anasayfa.dtResim.Recordset
            .AddNew
            .Fields("Siparis_Kodu") = CStr(i)
            .Fields("Resim_Kodu") = CStr(ResimNO)
            .Update
        End With
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''KAYIT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'sayfa temizlenir. �lk haline getirilir. yeni bir kay�t i�in yap�lan haz�rl�klar da denilebilir
    For Each Control In Me
        If TypeOf Control Is TextBox Then Control.Text = ""
    Next Control
    resim.Value = 0
    Timer1.Enabled = True   'ad� ve soyad� combolar�n� temizler ve adlar� comboya yazar
    'form kullan�ma a��l�r.
    Me.Enabled = True: Me.MousePointer = 1
    Write #7, Time$, Me.Name, "Command2_Click", "YeniM�steri:" & M�steriID 'logging
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+

Private Sub ad�_Change()
    Dim i As Long: Dim nSel As Long
    If GeriAl = True Or ad�.Text = "" Then GeriAl = False: Exit Sub
    For i = 0 To ad�.ListCount - 1
        If InStr(1, ad�.List(i), ad�.Text, vbTextCompare) = 1 Then
            nSel = ad�.SelStart: ad�.Text = ad�.List(i): ad�.SelStart = nSel: ad�.SelLength = Len(ad�.Text) - nSel
            Exit For
        End If
    Next
End Sub
Private Sub ad�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If ad�.Text <> "" Then GeriAl = True
    End If
End Sub
Private Sub ad�_LostFocus()
    On Local Error GoTo HataControl
    If Trim(ad�.Text) = "" Then Exit Sub
    Dim i As Integer
    soyad�.Clear: ad�.Text = UpperCaseFirstLetter(Trim(ad�.Text))
    Anasayfa.dtM�steri.Recordset.MoveFirst
    For i = 1 To Anasayfa.dtM�steri.Recordset.RecordCount
        If UCase(ad�.Text) = UCase(Anasayfa.dtM�steri.Recordset.Fields("Musteri_Adi")) Then soyad�.AddItem Anasayfa.dtM�steri.Recordset.Fields("Musteri_Soyadi")
        Anasayfa.dtM�steri.Recordset.MoveNext
    Next i
    Exit Sub
HataControl:
End Sub
Private Sub soyad�_Change()
    Dim i As Long: Dim nSel As Long
    If GeriAl = True Or soyad�.Text = "" Then GeriAl = False: Exit Sub
    For i = 0 To soyad�.ListCount - 1
        If InStr(1, soyad�.List(i), soyad�.Text, vbTextCompare) = 1 Then
            nSel = soyad�.SelStart: soyad�.Text = soyad�.List(i): soyad�.SelStart = nSel: soyad�.SelLength = Len(soyad�.Text) - nSel
            Exit For
        End If
    Next
End Sub
Private Sub soyad�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If soyad�.Text <> "" Then GeriAl = True
    End If
End Sub
Private Sub soyad�_LostFocus()
    On Local Error GoTo HataControl
    If Trim(soyad�.Text) = "" Then Exit Sub
    soyad�.Text = UpperCaseFirstLetter(Trim(soyad�.Text))
    Dim i As Integer: Dim M�steriID As String
    M�steriID = VarsaM�steriIDBul(ad�.Text, soyad�.Text)
    If M�steriID = "0" Then Exit Sub
    tel.Text = Anasayfa.dtM�steri.Recordset.Fields("Musteri_Telefon")
    With Anasayfa.dtSiparis.Recordset
        If .RecordCount <> 0 Then
            .MoveLast
            For i = 1 To .RecordCount
                If CStr(.Fields("Musteri_Kodu")) = CStr(M�steriID) Then
                    bel.Text = .Fields("Bel")
                    basen.Text = .Fields("Basen")
                    gogus.Text = .Fields("G���s")
                    omuz.Text = .Fields("Omuz")
                    kol.Text = .Fields("Kol")
                    boy.Text = .Fields("Boy")
                    Exit For
                End If
                .MovePrevious
            Next i
        End If
    End With
    Exit Sub
HataControl:
End Sub
Private Sub kumas_Change()
    Dim i As Long: Dim nSel As Long
    If GeriAl = True Or kumas.Text = "" Then GeriAl = False: Exit Sub
    For i = 0 To kumas.ListCount - 1
        If InStr(1, kumas.List(i), kumas.Text, vbTextCompare) = 1 Then
            nSel = kumas.SelStart: kumas.Text = kumas.List(i): kumas.SelStart = nSel: kumas.SelLength = Len(kumas.Text) - nSel
            Exit For
        End If
    Next
End Sub
Private Sub kumas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If kumas.Text <> "" Then GeriAl = True
    End If
End Sub
Private Sub kumas_LostFocus()
    kumas.Text = UpperCaseFirstLetter(Trim(kumas.Text))
End Sub
Private Sub cins_Change()
    Dim i As Long: Dim nSel As Long
    If GeriAl = True Or cins.Text = "" Then GeriAl = False: Exit Sub
    For i = 0 To cins.ListCount - 1
        If InStr(1, cins.List(i), cins.Text, vbTextCompare) = 1 Then
            nSel = cins.SelStart: cins.Text = cins.List(i): cins.SelStart = nSel: cins.SelLength = Len(cins.Text) - nSel
            Exit For
        End If
    Next
End Sub
Private Sub cins_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If cins.Text <> "" Then GeriAl = True
    End If
End Sub
Private Sub cins_LostFocus()
    cins.Text = UpperCaseFirstLetter(Trim(cins.Text))
End Sub
Private Sub tel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 32 Then Exit Sub
    If KeyAscii > 57 Or KeyAscii < 47 Then KeyAscii = 0
End Sub
''''''''''''''got focus komutlar�
Private Sub ad�_GotFocus(): Call SelectAllText: End Sub
Private Sub soyad�_GotFocus(): Call SelectAllText: End Sub
Private Sub tel_GotFocus(): Call SelectAllText: End Sub
Private Sub fiyat_GotFocus(): Call SelectAllText: End Sub
Private Sub kumas_GotFocus(): Call SelectAllText: End Sub
Private Sub cins_GotFocus(): Call SelectAllText: End Sub
Private Sub bel_GotFocus(): Call SelectAllText: End Sub
Private Sub basen_GotFocus(): Call SelectAllText: End Sub
Private Sub gogus_GotFocus(): Call SelectAllText: End Sub
Private Sub omuz_GotFocus(): Call SelectAllText: End Sub
Private Sub kol_GotFocus(): Call SelectAllText: End Sub
Private Sub boy_GotFocus(): Call SelectAllText: End Sub
Private Sub acik_GotFocus(): Call SelectAllText: End Sub
Private Sub fiyat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 13 Then Exit Sub
    If KeyAscii > 57 Or KeyAscii < 47 Then KeyAscii = 0
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub g�sterresim_Click()
    If Trim(ResAdres) = "" Then
        MsgBox "Yolun do�ru yaz�ld���ndan emin olun.", vbOKOnly, "Resim Bulunamad�."
    Else
        Shell (App.path + "\Resimci.exe " + ResAdres)
    End If
End Sub
Private Sub g�zat_Click()
    On Local Error GoTo HataControl
    CommonDialog1.Filter = "Ge�erli Resim Dosyalar�|*.jpg;*.jpeg;*.bmp;"
    CommonDialog1.ShowOpen
    ResAdres.Caption = CommonDialog1.FileName
    Exit Sub
HataControl:
End Sub
Private Sub olustur_Click()
    If Dir(App.path + "\Pictures\Dikimevi.jpg") <> "" Then Kill App.path + "\Pictures\Dikimevi.jpg"
    FileCopy App.path + "\pictures\sample.qaz", App.path + "\Pictures\Dikimevi.jpg"
    ResAdres.Caption = App.path + "\Pictures\DikimEvi.jpg"
    Shell "mspaint.exe """ + ResAdres.Caption + "", vbMaximizedFocus
End Sub
Private Sub resadres_Change()
    If Trim(ResAdres.Caption) <> "" Then g�sterresim.Enabled = True Else g�sterresim.Enabled = False
End Sub
Private Sub resim_Click()
    If resim.Value = 0 Then ResAdres.Caption = "": g�zat.Enabled = False: olustur.Enabled = False: g�sterresim.Enabled = False
    If resim.Value = 1 Then ResAdres.Caption = "": g�zat.Enabled = True: olustur.Enabled = True: g�zat.SetFocus
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+

Private Sub Timer1_Timer()
    Write #7, Time$, Me.Name, "Timer1_Timer", "Start" 'logging
    On Local Error Resume Next
    MousePointer = 11: Enabled = False
    Dim sayi As Integer: Dim i As Integer: Dim j As Integer: Dim M�steriName As String
    sip.Value = Date: pro.Value = Date: tes.Value = Date
    ad�.Clear: soyad�.Clear: cins.Clear: kumas.Clear
    'ilk olarak m�sterilerin adlar�n� comboya ekleyece�iz
    'ama listede zaten o ad varsa bir daha eklenmeyecek
    'bu y�zden eklemeden �nce t�m comboyu kontrol ediyoruz.
    With Anasayfa.dtM�steri.Recordset   'm�steriler
        sayi = .RecordCount
        .MoveFirst
        For i = 1 To sayi
            M�steriName = .Fields("Musteri_Adi")
            For j = 0 To ad�.ListCount
                If UCase(ad�.List(j)) = UCase(M�steriName) Then GoTo bas
            Next j
            ad�.AddItem M�steriName
bas:
            .MoveNext
        Next i
    End With
    With Anasayfa.dtKumasT�r�.Recordset   'kumas t�rleri
        .MoveFirst
        For i = 1 To .RecordCount
            kumas.AddItem .Fields("Kumas_Adi")
            .MoveNext
        Next i
    End With
    With Anasayfa.dtSiparisT�r�.Recordset   'siparis t�rleri (elbise cinsi)
        .MoveFirst
        For i = 1 To .RecordCount
            cins.AddItem .Fields("Siparis_Adi")
            .MoveNext
        Next i
    End With
    'comboya ad ekledikten sonra genel bilgi yaz�l�yor.
    kumas.ListIndex = 0: cins.ListIndex = 0
    kayno.Caption = "M�steri Say�s� : " & Anasayfa.dtM�steri.Recordset.RecordCount
    kayno2.Caption = "�s Say�s� : " & Anasayfa.dtSiparis.Recordset.RecordCount
    MousePointer = 1: Enabled = True
    Write #7, Time$, Me.Name, "Timer1_Timer", "Successful" 'logging
    Timer1.Enabled = False
End Sub
Private Sub Form_Load()
    Write #7, Time$, Me.Name, "Form_Load", "Start" 'logging
    Dim Control As Control: Me.BackColor = rnk_frm_arka: Timer1.Enabled = True 'ad� ve soyad� combolar�n� temizler ve adlar� comboya yazar
    For Each Control In Me
        If TypeOf Control Is OsenXPButton Then Control.BackColor = rnk_btn_arka: Control.BackOver = rnk_btn_arka: Control.ForeColor = rnk_btn_�n: Control.ForeOver = rnk_btn_�n
        If TypeOf Control Is Label Then Control.ForeColor = rnk_frm_�n
        If TypeOf Control Is Frame Then Control.BackColor = rnk_frm_arka: Control.ForeColor = rnk_frm_�n
        If TypeOf Control Is CheckBox Then Control.BackColor = rnk_frm_arka: Control.ForeColor = rnk_frm_�n
        If TypeOf Control Is ComboBox Then Control.BackColor = rnk_yaz�_arka: Control.ForeColor = rnk_yaz�_�n
        If TypeOf Control Is TextBox Then Control.BackColor = rnk_yaz�_arka: Control.ForeColor = rnk_yaz�_�n
        If TypeOf Control Is DTPicker Then Control.CalendarBackColor = rnk_yaz�_arka: Control.CalendarForeColor = rnk_yaz�_�n
        If TypeOf Control Is Line Then Control.BorderColor = rnk_frm_�n
    Next Control
    If Dir(Tema_Yeri & "\logo.gif") <> "" Then Image1.Picture = LoadPicture(Tema_Yeri & "\logo.gif")
    Me.Show: frmMain.Caption = App.ProductName + "-Siparis Olustur": Call frmMain.MDIForm_Resize
    Write #7, Time$, Me.Name, "Form_Load", "Successful" 'logging
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmMain.Caption = App.ProductName: Anasayfa.Visible = True: Anasayfa.Command1.SetFocus
    Write #7, Time$, Me.Name, "Form_Unload", "Successful" 'logging
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub


