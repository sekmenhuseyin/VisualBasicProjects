VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form Malii�lemler 
   BackColor       =   &H0096E06D&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7965
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMali.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   7965
   Begin OsenXPCntrl.OsenXPButton g�s 
      Default         =   -1  'True
      Height          =   735
      Left            =   5355
      TabIndex        =   0
      Top             =   7965
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "G�ster"
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
      MICON           =   "frmMali.frx":000C
      PICN            =   "frmMali.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer5 
      Interval        =   150
      Left            =   120
      Top             =   9120
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0096E06D&
      Caption         =   "��lem T�r�"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   195
      TabIndex        =   12
      Top             =   240
      Width           =   7575
      Begin MSComCtl2.MonthView viev 
         Height          =   2310
         Left            =   4605
         TabIndex        =   13
         Top             =   300
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   9887853
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   20905986
         CurrentDate     =   38559
      End
      Begin OsenXPCntrl.OsenXPButton Command1 
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "G�n� Ge�mi� Bor�"
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
         MICON           =   "frmMali.frx":047A
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton Command5 
         Height          =   495
         Left            =   2310
         TabIndex        =   5
         Top             =   300
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "G�nl�k Has�lat"
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
         MICON           =   "frmMali.frx":0496
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
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   915
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "T�m Bor�lar"
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
         MICON           =   "frmMali.frx":04B2
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton Command3 
         Height          =   495
         Left            =   2310
         TabIndex        =   6
         Top             =   945
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Ayl�k Has�lat"
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
         MICON           =   "frmMali.frx":04CE
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton Command8 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1515
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Para Giri�i"
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
         MICON           =   "frmMali.frx":04EA
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton Command4 
         Height          =   495
         Left            =   2310
         TabIndex        =   7
         Top             =   1515
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Y�ll�k Has�lat"
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
         MICON           =   "frmMali.frx":0506
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton Command7 
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   2100
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "M��teri Ara"
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
         MICON           =   "frmMali.frx":0522
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton Command6 
         Height          =   495
         Left            =   2310
         TabIndex        =   8
         Top             =   2100
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "T�m M��teriler"
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
         MICON           =   "frmMali.frx":053E
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0096E06D&
      Caption         =   "Bor� Listesi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   195
      TabIndex        =   11
      Top             =   3210
      Width           =   7575
      Begin VB.ListBox Liste 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3960
         ItemData        =   "frmMali.frx":055A
         Left            =   105
         List            =   "frmMali.frx":055C
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   5100
      End
      Begin VB.ListBox List1 
         Height          =   2205
         ItemData        =   "frmMali.frx":055E
         Left            =   210
         List            =   "frmMali.frx":0560
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2100
         Width           =   2430
      End
      Begin VB.ListBox List2 
         Height          =   2205
         ItemData        =   "frmMali.frx":0562
         Left            =   2625
         List            =   "frmMali.frx":0564
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2100
         Width           =   2430
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   4965
      End
      Begin VB.Label tarihlabel 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5340
         TabIndex        =   20
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tarih :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5280
         TabIndex        =   19
         Top             =   3360
         Width           =   570
      End
      Begin VB.Label tutar 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5340
         TabIndex        =   18
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tutar :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5280
         TabIndex        =   17
         Top             =   2025
         Width           =   585
      End
      Begin VB.Label i�lem 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   5340
         TabIndex        =   16
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yap�lan ��lem :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5280
         TabIndex        =   15
         Top             =   480
         Width           =   1260
      End
   End
   Begin OsenXPCntrl.OsenXPButton geri 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   195
      TabIndex        =   10
      Top             =   7995
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
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
      MICON           =   "frmMali.frx":0566
      PICN            =   "frmMali.frx":0582
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Labeltopbor� 
      Caption         =   "Label6"
      Height          =   540
      Left            =   3645
      TabIndex        =   24
      Top             =   4088
      Width           =   1170
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   375
      Left            =   4365
      TabIndex        =   21
      Top             =   4178
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label adet 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   14
      Top             =   7965
      Width           =   2625
   End
End
Attribute VB_Name = "Malii�lemler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TopBor� As String: Dim M��teri_Bilgisi As String
Dim Ad2 As String: Dim Soyad2 As String: Dim Bor�2 As String: Dim Tarih2 As String
Dim Ad As String: Dim Soyad As String: Dim Bor� As String: Dim tarih As String
Public Sub Command1_Click()
    Write #7, Time$, Me.Name, "Command1_Click", "Start" 'logging
    On Local Error Resume Next
    Dim G�n As String: Dim Ay As String: Dim Y�l As String: Dim i As Integer
    Me.Enabled = False: Me.MousePointer = 11
    Label2.Caption = "0": TopBor� = 0
    i�lem = "G�n� Ge�mi� Bor�lar"
    Label5.Caption = "Tarih :": tarihlabel = viev.Value
    Label3.Caption = "Toplam Bor�"
    Label4.Caption = "Ad        Soyad             Bor�"
    Liste.Clear: List1.Clear: List2.Clear
    Open App.path & "\Temp.txt" For Output As #1: Close #1
    Open App.path + "\temp2.txt" For Output As #1: Close #1
    Open App.path + "\temp3.txt" For Output As #1: Close #1
    'ilk �nce teslim tarihi ge�mi� olanlar� dosyaya yaz�lacak
    With Anasayfa.dtSipari�.Recordset
        .MoveFirst
        For i = 1 To .RecordCount
            Ad = .Fields("Musteri_Kodu"): Soyad = .Fields("Siparis_Kodu"): Bor� = .Fields("Ucret"): tarih = .Fields("Teslim_Tarihi")
            G�n = Left(tarih, 2): Ay = Mid(tarih, 4, 2): Y�l = Right(tarih, 4)
            If Val(Bor�) = 0 And CStr(Bor�) = "" Then GoTo ekleme
            If Val(Y�l) > Val(viev.Year) Then GoTo ekleme
            If Val(Y�l) = Val(viev.Year) And Val(Ay) > Val(viev.Month) Then GoTo ekleme
            If Val(Y�l) = Val(viev.Year) And Val(Ay) = Val(viev.Month) And Val(G�n) > Val(viev.Day) Then GoTo ekleme
            Open App.path + "\Temp.txt" For Append As #1: Write #1, tarih, Ad, Soyad, Bor�: Close #1
ekleme:
            .MoveNext
        Next i
        'ilk �nce teslim tarihi ge�mi� olanlar� dosyaya yaz�ld�
        '''''''''''''''''''''''''''''''''''''''
kontrolbas:
        Open App.path + "\temp2.txt" For Input As #1
        Open App.path & "\Temp.txt" For Input As #2
tekrar:
        If EOF(2) = True Then GoTo asd2
        Input #2, tarih, Ad, Soyad, Bor�
qwe2:
        If EOF(1) = True Then GoTo asd
        Input #1, Ad2, Soyad2, Bor�2
        If Ad = Ad2 Then
            Close #1 'daha �nce kaydedilmi� ad
            Open App.path + "\temp2.txt" For Input As #1
            GoTo tekrar
        Else
            GoTo qwe2 'kaydedilmemi�,      en az�ndan bu kay�ta e�it de�il
        End If
asd:
        Close #1    'hi� kaydedilmemi�.
dfg:        'burada her ki�inin toplam bor� hesaplan�yor.
        If EOF(2) = True Then GoTo yu�
        Input #2, Tarih2, Ad2, Soyad2, Bor�2
        If Ad = Ad2 Then Bor� = Val(Bor�) + Val(Bor�2)
        GoTo dfg
yu�:
        Close #2
        Open App.path + "\temp2.txt" For Append As #1
            Write #1, Ad, Soyad, Bor�
        Close #1
        Open App.path & "\Temp.txt" For Input As #2
        Open App.path + "\temp2.txt" For Input As #1
        GoTo tekrar
asd2:
        Close #1: Close #2
    End With
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''     borcu olan t�m m��terilerin toplam �demesi gereken miktar bulundu
    '''     �imdi de �dedikleri miktarlar hesaplanacak...
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With Anasayfa.dt�deme.Recordset
        Open App.path + "\temp2.txt" For Input As #1
q1:
        If EOF(1) = True Then GoTo q4
        Input #1, Ad, Soyad, Bor� 'Musteri_Kodu,Siparis_Kodu,toplam �demesi gereken miktar
        Bor�2 = "0": .MoveFirst
        For i = 1 To .RecordCount
            If Val(.Fields("Musteri_Kodu")) = Val(Ad) Then Bor�2 = Val(Bor�2) + Val(.Fields("Odenen_Fiyat"))
            .MoveNext
        Next i
        Bor� = CStr(Val(Bor�) - Val(Bor�2))
        If Bor� <> "0" Then
            Open App.path + "\temp3.txt" For Append As #2
                Write #2, Ad, Bor�
            Close #2
            List2.AddItem Ad
            Ad = GetM��terifromID(Ad)
            TopBor� = Val(Bor�) + Val(TopBor�)
            Liste.AddItem Ad & "   " & Bor�
            List1.AddItem Ad & "   " & Bor�
        End If
        GoTo q1
q4:
        Close #1
    End With
    If Liste.ListCount >= 1 Then Liste.Selected(0) = True: adet.Caption = Liste.ListCount & " Kay�t Listede G�r�nt�lendi." Else adet.Caption = "G�r�nt�lenecek Kay�t Yok."
    If TopBor� <> 0 Then tutar.Caption = TopBor� & " YTL" Else tutar.Caption = "0 YTL"
    Labeltopbor�.Caption = TopBor�
    Me.Enabled = True: Me.MousePointer = 1
    Liste.SetFocus
    Write #7, Time$, Me.Name, "Command1_Click", "Successful" 'logging
End Sub
Public Sub Command2_Click()
    Write #7, Time$, Me.Name, "Command2_Click", "Start" 'logging
    On Local Error Resume Next
    Dim i As Integer
    Me.Enabled = False: Me.MousePointer = 11
    Label2.Caption = "0": TopBor� = 0
    i�lem = "Toplam Bor�"
    Label5.Caption = "Tarih :": tarihlabel = viev.Value
    Label3.Caption = "Toplam Bor�"
    Label4.Caption = "Ad        Soyad             Bor�"
    Liste.Clear: List1.Clear: List2.Clear
    Open App.path & "\Temp.txt" For Output As #1: Close #1
    Open App.path + "\temp2.txt" For Output As #2: Close #2
    Open App.path + "\temp3.txt" For Output As #3: Close #3
    'ilk �nce teslim tarihi ge�mi� olanlar� dosyaya yaz�lacak
    With Anasayfa.dtSipari�.Recordset
        .MoveFirst
        For i = 1 To .RecordCount
            Ad = .Fields("Musteri_Kodu"): Soyad = .Fields("Siparis_Kodu"): Bor� = .Fields("Ucret")
            If Val(Bor�) <> 0 And CStr(Bor�) <> "" Then
                Open App.path + "\Temp.txt" For Append As #1: Write #1, Ad, Soyad, Bor�: Close #1
            End If
            .MoveNext
        Next i
        'ilk �nce teslim tarihi ge�mi� olanlar� dosyaya yaz�ld�
        '''''''''''''''''''''''''''''''''''''''
kontrolbas:
        Open App.path + "\temp2.txt" For Input As #2
        Open App.path & "\Temp.txt" For Input As #1
tekrar:
        If EOF(1) = True Then GoTo asd2
        Input #1, Ad, Soyad, Bor�       'ilk olarak bu isim 2. dosyada varm� bak�l�r
qwe2:                                   'e�er varsa 1. dosyadan bir sonraki ad i�in ayn� i�lem yap�l�r.
        If EOF(2) = True Then GoTo asd  'yoksa 1. dosyadaki o isme ait toplam bor� hesaplan�r.
        Input #2, Ad2, Soyad2, Bor�2    't�m isimlerden sonra ise o ki�ilerin �dedi�i miktarlar bulunup
        If Ad = Ad2 Then                'toplam bor�lar�ndan ��kar�l�r.
            Close #2 'daha �nce kaydedilmi� ad
            Open App.path + "\temp2.txt" For Input As #2
            GoTo tekrar
        Else
            GoTo qwe2 'kaydedilmemi�,      en az�ndan bu kay�ta e�it de�il
        End If
asd:
        Close #2    'hi� kaydedilmemi�.
dfg:        'burada her ki�inin toplam bor� hesaplan�yor.
        If EOF(1) = True Then GoTo yu�
        Input #1, Ad2, Soyad2, Bor�2
        If Ad = Ad2 Then Bor� = Val(Bor�) + Val(Bor�2)
        GoTo dfg
yu�:
        Close #1
        Open App.path + "\temp2.txt" For Append As #2
            Write #2, Ad, Soyad, Bor�
        Close #2
        Open App.path & "\Temp.txt" For Input As #1
        Open App.path + "\temp2.txt" For Input As #2
        GoTo tekrar
asd2:
        Close #1: Close #2
    End With
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''     borcu olan t�m m��terilerin toplam �demesi gereken miktar bulundu
    '''     �imdi de �dedikleri miktarlar hesaplanacak...
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Open App.path + "\temp2.txt" For Input As #2
    With Anasayfa.dt�deme.Recordset
q1:
        If EOF(2) = True Then GoTo q4
        Input #2, Ad, Soyad, Bor� 'Musteri_Kodu,Siparis_Kodu,toplam �demesi gereken miktar
        Bor�2 = "0": .MoveFirst
        For i = 1 To .RecordCount
            If Val(.Fields("Musteri_Kodu")) = Val(Ad) Then Bor�2 = Val(Bor�2) + Val(.Fields("Odenen_Fiyat"))
            .MoveNext
        Next i
        Bor� = CStr(Val(Bor�) - Val(Bor�2))
        If Bor� <> "0" Then
            Open App.path + "\temp3.txt" For Append As #1
                Write #1, Ad, Bor�
            Close #1
            List2.AddItem Ad
            Ad = GetM��terifromID(Ad)
            TopBor� = Val(Bor�) + Val(TopBor�)
            Liste.AddItem Ad & "   " & Bor�
            List1.AddItem Ad & "   " & Bor�
        End If
        GoTo q1
q4:
        Close #2
    End With
    If Liste.ListCount >= 1 Then Liste.Selected(0) = True: adet.Caption = Liste.ListCount & " Kay�t Listede G�r�nt�lendi." Else adet.Caption = "G�r�nt�lenecek Kay�t Yok."
    If TopBor� <> 0 Then tutar.Caption = TopBor� & " YTL" Else tutar.Caption = "0 YTL"
    Labeltopbor�.Caption = TopBor�
    Me.Enabled = True: Me.MousePointer = 1
    Liste.SetFocus
    Write #7, Time$, Me.Name, "Command2_Click", "Successful" 'logging
End Sub
Public Sub Command3_Click()
    Write #7, Time$, Me.Name, "Command3_Click", "Start" 'logging
    On Local Error Resume Next
    Me.Enabled = False: Me.MousePointer = 11
    Label2.Caption = "0"
    Label4.Caption = "Ad        Soyad             Tutar"
    tutar.Caption = "0": Label3.Caption = "Tutar"
    Liste.Clear: i�lem.Caption = "Ayl�k Has�lat"
    Liste.Clear: List1.Clear: List2.Clear
    Label5.Caption = "Tarih :": tarihlabel = Right(viev.Value, 7)
    Open App.path + "\temp3.txt" For Output As #1: Close #1
    Open App.path + "\temp.txt" For Output As #1
    With Anasayfa.dtSipari�.Recordset
        .MoveFirst
bas:
        If .EOF() = True Then GoTo son
        If CStr(Right(.Fields("Teslim_Tarihi"), 7)) = CStr(Right(viev.Value, 7)) Then
            If CStr(.Fields("Ucret")) <> "" And Val(.Fields("Ucret")) <> 0 Then
                tutar.Caption = Val(tutar.Caption) + Val(.Fields("Ucret"))
                Write #1, .Fields("Teslim_Tarihi"), .Fields("Musteri_Kodu"), .Fields("Ucret")
            End If
        End If
        .MoveNext
        GoTo bas
son:
        Close #1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
kontrolbas:
Open App.path + "\temp3.txt" For Input As #1: Open App.path & "\temp.txt" For Input As #2

tekrar:
  If EOF(2) = True Then GoTo asd2
  Input #2, tarih, Ad, Bor�                                 'bu kay�t var m� diye bak�caz �imdi
qwe2:
    If EOF(1) = True Then GoTo asd
    Input #1, Ad2, Bor�2                                    'yukar�daki kay�t buna e�it mi diyoruz.

      If Ad = Ad2 Then
        'daha �nce kaydedilmi� ad
        Close #1
        Open App.path + "\temp3.txt" For Input As #1        'yani e�it,     daha �nce kaydedilmi�
        GoTo tekrar
      Else
        'kaydedilmemi�
        GoTo qwe2                                           'daha de�il. belki bir sonrakinde
      End If
asd:
    Close #1: Bor� = "0" 'hi� kaydedilmemi�.
    Close #2: Open App.path & "\temp.txt" For Input As #2
dfg:
    If EOF(2) = True Then GoTo yu�
    Input #2, Tarih2, Ad2, Bor�2
    If Ad = Ad2 Then Bor� = Val(Bor�) + Val(Bor�2)
    GoTo dfg
yu�:
    Close #2
''''''''''''''
    Open App.path + "\temp3.txt" For Append As #1
        Write #1, Ad, Bor�
    Close #1
    List2.AddItem Ad
    Ad = GetM��terifromID(Ad)
    Liste.AddItem Ad & "   " & Bor�
    List1.AddItem Ad & "   " & Bor�
''''''''''''''
    Open App.path & "\temp.txt" For Input As #2:    Open App.path + "\temp3.txt" For Input As #1
    GoTo tekrar
asd2:
  Close #1: Close #2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If Liste.ListCount >= 1 Then Liste.Selected(0) = True: adet.Caption = Liste.ListCount & " Kay�t Listede G�r�nt�lendi." Else adet.Caption = "G�r�nt�lenecek Kay�t Yok."
  If tutar.Caption <> "0" Then tutar.Caption = tutar.Caption & " YTL" Else tutar.Caption = "0 YTL"
  End With
  Labeltopbor�.Caption = tutar.Caption
  Me.Enabled = True: Me.MousePointer = 1
  Liste.SetFocus
    Write #7, Time$, Me.Name, "Command3_Click", "Successful" 'logging
End Sub
Public Sub Command4_Click()
    Write #7, Time$, Me.Name, "Command4_Click", "Start" 'logging
    On Local Error Resume Next
    Me.Enabled = False: Me.MousePointer = 11
    Label2.Caption = "0"
    Label4.Caption = "Ad        Soyad             Tutar"
    tutar.Caption = "0": Label3.Caption = "Tutar"
    Liste.Clear: i�lem.Caption = "Y�ll�k Has�lat"
    Liste.Clear: List1.Clear: List2.Clear
    Label5.Caption = "Tarih :": tarihlabel = Right(viev.Value, 4)
    Open App.path + "\temp3.txt" For Output As #1: Close #1
    Open App.path + "\temp.txt" For Output As #1
    With Anasayfa.dtSipari�.Recordset
        .MoveFirst
bas:
        If .EOF() = True Then GoTo son
        If CStr(Right(.Fields("Teslim_Tarihi"), 4)) = CStr(Right(viev.Value, 4)) Then
            If CStr(.Fields("Ucret")) <> "" And Val(.Fields("Ucret")) <> 0 Then
                tutar.Caption = Val(tutar.Caption) + Val(.Fields("Ucret"))
                Write #1, .Fields("Teslim_Tarihi"), .Fields("Musteri_Kodu"), .Fields("Ucret")
            End If
        End If
        .MoveNext
        GoTo bas
son:
        Close #1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
kontrolbas:
Open App.path + "\temp3.txt" For Input As #1
Open App.path & "\temp.txt" For Input As #2

tekrar:
  If EOF(2) = True Then GoTo asd2
  Input #2, tarih, Ad, Bor�                            'bu kay�t var m� diye bak�caz �imdi
qwe2:
    If EOF(1) = True Then GoTo asd
    Input #1, Ad2, Bor�2                                'yukar�daki kay�t buna e�it mi diyoruz.

      If Ad = Ad2 Then
        'daha �nce kaydedilmi� ad
        Close #1
        Open App.path + "\temp3.txt" For Input As #1        'yani e�it,     daha �nce kaydedilmi�
        GoTo tekrar
      Else
        'kaydedilmemi�
        GoTo qwe2                                           'daha de�il. belki bir sonrakinde
      End If
asd:
    Close #1: Bor� = "0" 'hi� kaydedilmemi�.
    Close #2: Open App.path & "\temp.txt" For Input As #2
dfg:
    If EOF(2) = True Then GoTo yu�
    Input #2, Tarih2, Ad2, Bor�2
    If Ad = Ad2 Then Bor� = Val(Bor�) + Val(Bor�2)
    GoTo dfg
yu�:
    Close #2
''''''''''''''
    Open App.path + "\temp3.txt" For Append As #1
        Write #1, Ad, Bor�
    Close #1
    List2.AddItem Ad
    Ad = GetM��terifromID(Ad)
    Liste.AddItem Ad & "   " & Bor�
    List1.AddItem Ad & "   " & Bor�
''''''''''''''
    Open App.path & "\temp.txt" For Input As #2: Open App.path + "\temp3.txt" For Input As #1
    GoTo tekrar
asd2:
  Close #1: Close #2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If Liste.ListCount >= 1 Then Liste.Selected(0) = True: adet.Caption = Liste.ListCount & " Kay�t Listede G�r�nt�lendi." Else adet.Caption = "G�r�nt�lenecek Kay�t Yok."
  If tutar.Caption <> "0" Then tutar.Caption = tutar.Caption & " YTL" Else tutar.Caption = "0 YTL"
  End With
  Labeltopbor�.Caption = tutar.Caption
  Me.Enabled = True: Me.MousePointer = 1
  Liste.SetFocus
    Write #7, Time$, Me.Name, "Command4_Click", "Successful" 'logging
End Sub
Public Sub Command5_Click()
    Write #7, Time$, Me.Name, "Command5_Click", "Start" 'logging
    On Local Error Resume Next
    Me.Enabled = False: Me.MousePointer = 11
    Label2.Caption = "0"
    Label4.Caption = "Ad        Soyad             Tutar"
    tutar.Caption = "0": Label3.Caption = "Tutar"
    i�lem.Caption = "G�nl�k Has�lat"
    Label5.Caption = "Tarih :": tarihlabel = viev.Value
    Liste.Clear: List1.Clear: List2.Clear
    Open App.path + "\temp3.txt" For Output As #1: Close #1
    Open App.path + "\temp.txt" For Output As #2
    With Anasayfa.dtSipari�.Recordset
        .MoveFirst
bas:
        If .EOF() = True Then GoTo son
        If CStr(.Fields("Teslim_Tarihi")) = CStr(viev.Value) Then
            If CStr(.Fields("Ucret")) <> "" And Val(.Fields("Ucret")) <> 0 Then
                tutar.Caption = Val(tutar.Caption) + Val(.Fields("Ucret"))
                Write #2, .Fields("Teslim_Tarihi"), .Fields("Musteri_Kodu"), .Fields("Ucret")
            End If
        End If
        .MoveNext
        GoTo bas
son:
        Close #2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
kontrolbas:
Open App.path + "\temp3.txt" For Input As #1: Open App.path & "\temp.txt" For Input As #2
tekrar:
  If EOF(2) = True Then GoTo asd2
  Input #2, tarih, Ad, Bor�                                 'bu kay�t var m� diye bak�caz �imdi
qwe2:
    If EOF(1) = True Then GoTo asd
    Input #1, Ad2, Bor�2                                    'yukar�daki kay�t buna e�it mi diyoruz.
    If Ad = Ad2 Then
        'daha �nce kaydedilmi� ad
        Close #1
        Open App.path + "\temp3.txt" For Input As #1        'yani e�it,     daha �nce kaydedilmi�
        GoTo tekrar
    Else
        'kaydedilmemi�
        GoTo qwe2                                           'daha de�il. belki bir sonrakinde
    End If
asd:
    Close #1: Bor� = 0 'hi� kaydedilmemi�.
    Close #2: Open App.path & "\temp.txt" For Input As #2
dfg:
    If EOF(2) = True Then GoTo yu�
    Input #2, Tarih2, Ad2, Bor�2
    If Ad = Ad2 Then Bor� = Val(Bor�) + Val(Bor�2)
    GoTo dfg
yu�:
    Close #2
''''''''''''''
    Open App.path + "\temp3.txt" For Append As #1
        Write #1, Ad, Bor�
    Close #1
    List2.AddItem Ad
    Ad = GetM��terifromID(Ad)
    Liste.AddItem Ad & "   " & Bor�
    List1.AddItem Ad & "   " & Bor�
''''''''''''''
    Open App.path & "\temp.txt" For Input As #2:    Open App.path + "\temp3.txt" For Input As #1
    GoTo tekrar
asd2:
  Close #1: Close #2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If Liste.ListCount >= 1 Then Liste.Selected(0) = True: adet.Caption = Liste.ListCount & " Kay�t Listede G�r�nt�lendi." Else adet.Caption = "G�r�nt�lenecek Kay�t Yok."
  If tutar.Caption <> "0" Then tutar.Caption = tutar.Caption & " YTL" Else tutar.Caption = "0 YTL"
  End With
  Labeltopbor�.Caption = tutar.Caption
  Me.Enabled = True: Me.MousePointer = 1
  Liste.SetFocus
    Write #7, Time$, Me.Name, "Command5_Click", "Successful" 'logging
End Sub
Public Sub Command6_Click()
    Write #7, Time$, Me.Name, "Command6_Click", "Start" 'logging
    On Local Error Resume Next
    Me.Enabled = False: Me.MousePointer = 11
    Label2.Caption = "0": TopBor� = 0
    Liste.Clear: List1.Clear: List2.Clear: tarihlabel = Date
    i�lem = "T�m M��teriler": Label5.Caption = "Tarih :"
    Label3.Caption = "Tutar": Label4.Caption = "Ad        Soyad             Tutar"
    Open App.path & "\Temp.txt" For Output As #1: Close #1
    Open App.path + "\temp2.txt" For Output As #1: Close #1
    Open App.path + "\temp3.txt" For Output As #1: Close #1
    With Anasayfa.dtSipari�.Recordset
        .MoveFirst
bas:
        If .EOF() = True Then GoTo son
        Ad = .Fields("Musteri_Kodu")
        Bor� = .Fields("Ucret")
        If Bor� <> "" Then
        Else
            GoTo ekleme
        End If
        Open App.path + "\Temp.txt" For Append As #1: Write #1, tarih, Ad, Bor�: Close #1
ekleme:
        .MoveNext
        GoTo bas
son:
'''''''''''''''''''''''''''''''''''''''
kontrolbas:
        Open App.path + "\temp2.txt" For Input As #1
        Open App.path & "\Temp.txt" For Input As #2
tekrar:
        If EOF(2) = True Then GoTo asd2
        Input #2, tarih, Ad, Bor�
qwe2:
        If EOF(1) = True Then GoTo asd
        Input #1, Ad2, Bor�2
        If Ad = Ad2 Then
            'daha �nce kaydedilmi� ad
            Close #1
            Open App.path + "\temp2.txt" For Input As #1
            GoTo tekrar
        Else
            'kaydedilmemi�
            GoTo qwe2
        End If

asd:
        Close #1    'hi� kaydedilmemi�.
dfg:
        If EOF(2) = True Then GoTo yu�
        Input #2, Tarih2, Ad2, Bor�2
        If Ad = Ad2 Then Bor� = Val(Bor�) + Val(Bor�2)
        GoTo dfg
yu�:
        Close #2
        Open App.path + "\temp2.txt" For Append As #1
            Write #1, Ad, Bor�
        Close #1
        ''''''''''''''
        Open App.path + "\temp3.txt" For Append As #1
            Write #1, Ad, Bor�
        Close #1
        TopBor� = Val(Bor�) + Val(TopBor�)
        List2.AddItem Ad
        Ad = GetM��terifromID(Ad)
        Liste.AddItem Ad & "   " & Bor�
        List1.AddItem Ad & "   " & Bor�
        ''''''''''''''
        Open App.path & "\Temp.txt" For Input As #2
        Open App.path + "\temp2.txt" For Input As #1
        GoTo tekrar
asd2:
    End With
    Close #1: Close #2
    If Liste.ListCount >= 1 Then Liste.Selected(0) = True: adet.Caption = Liste.ListCount & " Kay�t Listede G�r�nt�lendi." Else adet.Caption = "G�r�nt�lenecek Kay�t Yok."
    If TopBor� <> "0" Then tutar.Caption = TopBor� & " YTL" Else tutar.Caption = "0 YTL"
    Labeltopbor�.Caption = TopBor�
    Me.Enabled = True: Me.MousePointer = 1
    Liste.SetFocus
    Write #7, Time$, Me.Name, "Command6_Click", "Successful" 'logging
End Sub
Public Sub Command7_Click()
    Write #7, Time$, Me.Name, "Command7_Click", "Start" 'logging
    Dim temp As String: Dim tmp As Integer
    M��teri_Bilgisi = InputBox("Aramak istedi�iniz m��terinin herhangi bir bilgisini giriniz.", "M��teri Arama")
    Open App.path + "\temp3.txt" For Output As #1: Close #1
    Open App.path + "\temp.txt" For Output As #1: Close #1
    Liste.Clear: List1.Clear: List2.Clear
Start:
    M��teri_Bilgisi = Trim(M��teri_Bilgisi)
    If M��teri_Bilgisi = "" Then Exit Sub
    tmp = InStr(1, M��teri_Bilgisi, " ")
    If tmp <> 0 Then
        temp = Left(M��teri_Bilgisi, tmp)
        M��teri_Bilgisi = Right(M��teri_Bilgisi, Len(M��teri_Bilgisi) - tmp)
        Arama_��lemi (temp)
        GoTo Start
    End If
    Arama_��lemi (M��teri_Bilgisi)
    Kay�t��inSay�m 2: Label2.Caption = "0"
    Write #7, Time$, Me.Name, "Command7_Click", "Successful" 'logging
End Sub
Public Sub Arama_��lemi(Kriter As String)
    On Local Error Resume Next
    Me.Enabled = False: Me.MousePointer = 11
    Label3.Caption = "Aranan": tutar.Caption = """" & M��teri_Bilgisi & """"
    Label4.Caption = "Ad        Soyad    (Kay�t No)"
    Label5.Caption = "": tarihlabel.Caption = ""
    i�lem = "M��teri Arama"
    'arama i�lemleri a�a��dad�r...
    Arama_ve_Ekleme "Musteri_Adi", CStr(M��teri_Bilgisi)
    Arama_ve_Ekleme "Musteri_Soyadi", CStr(M��teri_Bilgisi)
    If List1.ListCount >= 1 Then Liste.Selected(0) = True: adet.Caption = List1.ListCount & " Kay�t Listede G�r�nt�lendi." Else adet.Caption = "G�r�nt�lenecek Kay�t Yok."
    Anasayfa.dtM��teri.RecordSource = "tbl_Musteriler": Anasayfa.dtM��teri.Refresh
    Me.Enabled = True: Me.MousePointer = 1
    Liste.SetFocus
End Sub
Sub Arama_ve_Ekleme(Alan As String, ifade As String)
    Write #7, Time$, Me.Name, "Arama_ve_Ekleme", "Start" 'logging
    On Local Error Resume Next
    Dim Kimlik, Arama_Kriteri As String: Dim �nceden_Yaz�lm�� As Boolean: Dim i As Integer
    Arama_Kriteri = "[" + Alan + "] like ""*" + Trim(ifade) + "*"""
    Anasayfa.dtM��teri.RecordSource = "tbl_Musteriler": Anasayfa.dtM��teri.Refresh
    Arama_Kriteri = "select * from tbl_Musteriler where " + Arama_Kriteri: Anasayfa.dtM��teri.RecordSource = Arama_Kriteri: Anasayfa.dtM��teri.Refresh
    Open App.path + "\temp3.txt" For Append As #2
    With Anasayfa.dtM��teri.Recordset
        Kay�t��inSay�m 2: If .RecordCount = 0 Then GoTo bitti
        For i = 1 To .RecordCount
            �nceden_Yaz�lm�� = False
            Open App.path + "\temp.txt" For Input As #1
            Do While EOF(1) = False
                Input #1, Kimlik: If CStr(Kimlik) = CStr(.Fields("Musteri_Kodu")) Then �nceden_Yaz�lm�� = True
            Loop
            Close #1
            If �nceden_Yaz�lm�� = False Then
                Open App.path + "\temp.txt" For Append As #1: Write #1, .Fields("Musteri_Kodu"): Close #1
                List2.AddItem .Fields("Musteri_Kodu")
                List1.AddItem .Fields("Musteri_Adi") & " " & .Fields("Musteri_Soyadi") & "     (" & .Fields("Musteri_Kodu") & ")"
                Liste.AddItem .Fields("Musteri_Adi") & " " & .Fields("Musteri_Soyadi") & "     (" & .Fields("Musteri_Kodu") & ")"
            End If
            Write #2, .Fields("Musteri_Kodu"), "0"
            .MoveNext
        Next i
bitti:
        Close #2
    End With
    Anasayfa.dtM��teri.RecordSource = "tbl_Musteriler": Anasayfa.dtM��teri.Refresh
    Write #7, Time$, Me.Name, "Arama_ve_Ekleme", Alan & "=" & ifade 'logging
End Sub
Private Sub Command8_Click()
    ParaGiri�i.Show: Me.Enabled = False: ParaGiri�i.startingpoint.Caption = Me.Name
End Sub
Private Sub Form_Load()
    Write #7, Time$, Me.Name, "Form_Load", "Start" 'logging
    Dim Control As Control: Me.BackColor = rnk_frm_arka
    viev.Value = Date: viev.MonthBackColor = rnk_yaz�_arka: viev.BackColor = rnk_frm_arka: viev.TitleBackColor = rnk_frm_arka: viev.TitleForeColor = rnk_yaz�_�n: viev.TrailingForeColor = rnk_yaz�_�n: viev.ForeColor = rnk_yaz�_�n
    For Each Control In Me
        If TypeOf Control Is OsenXPButton Then Control.BackColor = rnk_btn_arka: Control.BackOver = rnk_btn_arka: Control.ForeColor = rnk_btn_�n: Control.ForeOver = rnk_btn_�n
        If TypeOf Control Is Label Then Control.ForeColor = rnk_frm_�n
        If TypeOf Control Is ListBox Then Control.ForeColor = rnk_yaz�_�n: Control.BackColor = rnk_yaz�_arka
        If TypeOf Control Is Frame Then Control.BackColor = rnk_frm_arka: Control.ForeColor = rnk_frm_�n
    Next Control
    Me.Show: frmMain.Caption = App.ProductName + "-Mali ��lemler": Call frmMain.MDIForm_Resize
    Write #7, Time$, Me.Name, "Form_Load", "Successful" 'logging
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmMain.Caption = App.ProductName: Anasayfa.Visible = True: Anasayfa.Command1.SetFocus
    Write #7, Time$, Me.Name, "Form_Unload", "Successful" 'logging
End Sub
Public Sub geri_Click()
    Dim i As Integer
    If Label2.Caption = "1" Then
      Me.Enabled = False: Me.MousePointer = 11
      Liste.Clear: Label2.Caption = "0"
      For i = 0 To List1.ListCount - 1:
        Liste.AddItem List1.List(i)
      Next i
      If Liste.ListCount >= 1 Then Liste.Selected(0) = True: adet.Caption = Liste.ListCount & " Kay�t Listede G�r�nt�lendi." Else adet.Caption = "G�r�nt�lenecek Kay�t Yok."
      Select Case i�lem.Caption
      Case "G�n� Ge�mi� Bor�lar"
        Label3.Caption = "Toplam Bor�"
        Label4.Caption = "Ad        Soyad             Bor�"
        tutar.Caption = Labeltopbor�.Caption
        Label5.Caption = "Tarih :": tarihlabel.Caption = Date
      Case "T�m Bor�lar"
        Label3.Caption = "Toplam Bor�"
        Label4.Caption = "Ad        Soyad             Bor�"
        tutar.Caption = Labeltopbor�.Caption
        Label5.Caption = "Tarih :": tarihlabel.Caption = Date
      Case "T�m M��teriler"
        Label3.Caption = "Tutar"
        Label4.Caption = "Ad        Soyad             Tutar"
        tutar.Caption = Labeltopbor�.Caption
        Label5.Caption = "Tarih :": tarihlabel.Caption = Date
      Case "M��teri Arama"
        Label3.Caption = "Aranan"
        Label4.Caption = "Ad        Soyad"
        tutar.Caption = """" & M��teri_Bilgisi & """"
        Label5.Caption = "Toplam Borcu :": tarihlabel.Caption = TopBor�
      Case "G�nl�k Has�lat"
        Label3.Caption = "Tutar"
        Label4.Caption = "Ad        Soyad             Tutar"
        tutar.Caption = Labeltopbor�.Caption
        Label5.Caption = "Tarih :": tarihlabel.Caption = viev.Value
      Case "Ayl�k Has�lat"
        Label3.Caption = "Tutar"
        Label4.Caption = "Ad        Soyad             Tutar"
        tutar.Caption = Labeltopbor�.Caption
         Label5.Caption = "Tarih :": tarihlabel.Caption = Right(viev.Value, 7)
     Case "Y�ll�k Has�lat"
        Label3.Caption = "Tutar"
        Label4.Caption = "Ad        Soyad             Tutar"
        tutar.Caption = Labeltopbor�.Caption
         Label5.Caption = "Tarih :": tarihlabel.Caption = Right(viev.Value, 4)
     End Select
      Me.Enabled = True: Me.MousePointer = 1
      Liste.SetFocus
    Else
      Anasayfa.Visible = True: Unload Me
    End If
End Sub
Private Sub g�s_Click()
    If Trim(Liste.Text = "") Then Exit Sub
    If Label2.Caption = "0" Then
        Liste_DblClick
    Else
        Dim tmp As String
        Open App.path + "\temp2.txt" For Input As #2: Input #2, tmp: Close #2
        Open App.path + "\temp2.txt" For Output As #2: Write #2, tmp, ListeNO_Bul(Liste.Text): Close #2
        Me.Visible = False: G�ster.Show: G�ster.startingpoint.Caption = Me.Name
    End If
End Sub
Private Sub Liste_DblClick()
    Write #7, Time$, Me.Name, "Liste_DblClick", "Start" 'logging
    On Local Error Resume Next
    Dim temp As String: Dim i, j As Integer
    Me.Enabled = False: Me.MousePointer = 11
    Label5.Caption = "Tarih :": tarihlabel.Caption = Date: Label3.Caption = "Tutar"
    If Label2.Caption = "0" Then Label2.Caption = "1" Else g�s_Click: Exit Sub
Devam:
    For i = 0 To Liste.ListCount
        If CStr(Liste.Text) = CStr(List1.List(i)) Then Exit For
    Next i
    Ad = List2.List(i): Liste.Clear: TopBor� = 0
    With Anasayfa.dtSipari�.Recordset
        If .RecordCount = 0 Then GoTo bitti
        .MoveFirst
        For i = 1 To .RecordCount
            If Val(Ad) = Val(.Fields("Musteri_Kodu")) Then
                Anasayfa.dtSipari�T�r�.Recordset.MoveFirst
                For j = 1 To Anasayfa.dtSipari�T�r�.Recordset.RecordCount
                    If Val(Anasayfa.dtSipari�T�r�.Recordset.Fields("Siparis_Turleri")) = Val(.Fields("Siparis_Turu")) Then temp = Anasayfa.dtSipari�T�r�.Recordset.Fields("Siparis_Adi"): Exit For
                    Anasayfa.dtSipari�T�r�.Recordset.MoveNext
                Next j
                Liste.AddItem "(" & .Fields("Durum") & ") " & .Fields("Teslim_Tarihi") & " " & GetM��terifromID(Ad) & " " & temp & " (" & .Fields("Siparis_Kodu") & ")"
                TopBor� = Val(TopBor�) + Val(.Fields("Ucret"))
                Open App.path + "\temp2.txt" For Output As #2: Write #2, .Fields("Musteri_Kodu"): Close #2
            End If
            .MoveNext
        Next i
bitti:
    End With
    tutar.Caption = TopBor� & " YTL"
    If Liste.ListCount >= 1 Then Liste.Selected(0) = True: adet.Caption = Liste.ListCount & " Kay�t Listede G�r�nt�lendi." Else adet.Caption = "G�r�nt�lenecek Kay�t Yok."
    Label4.Caption = "Durum           Teslim Tarihi        Ad Soyad                  Cinsi         Kay�t No"
    Me.Enabled = True: Me.MousePointer = 1
    Liste.SetFocus
    Write #7, Time$, Me.Name, "Liste_DblClick", "Successful" 'logging
End Sub
Private Sub Timer5_Timer()
    On Local Error Resume Next
    Command5_Click
    g�s.SetFocus
    Timer5.Enabled = False
End Sub
Private Sub viev_Click()
    Timer5.Enabled = True
End Sub

