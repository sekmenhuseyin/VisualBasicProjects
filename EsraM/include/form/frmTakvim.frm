VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form Takvim 
   BackColor       =   &H0096E06D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Takvim"
   ClientHeight    =   3165
   ClientLeft      =   6000
   ClientTop       =   3030
   ClientWidth     =   5640
   Icon            =   "frmTakvim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin OsenXPCntrl.OsenXPButton Command2 
      Height          =   735
      Left            =   3000
      TabIndex        =   5
      Top             =   105
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Sipari� Tarihlerini G�ster"
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
      MICON           =   "frmTakvim.frx":0442
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
      Height          =   735
      Left            =   3000
      TabIndex        =   4
      Top             =   840
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Prova Tarihlerini G�ster"
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
      MICON           =   "frmTakvim.frx":045E
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
      Height          =   735
      Left            =   3000
      TabIndex        =   3
      Top             =   1575
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Teslim Tarihlerini G�ster"
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
      MICON           =   "frmTakvim.frx":047A
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
      Height          =   735
      Left            =   3000
      TabIndex        =   2
      Top             =   2310
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "T�m Yap�lacaklar� G�ster"
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
      MICON           =   "frmTakvim.frx":0496
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton Command1 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   495
      Left            =   105
      TabIndex        =   1
      Top             =   2550
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Kabul Et"
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
      MICON           =   "frmTakvim.frx":04B2
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.MonthView viev 
      Height          =   2370
      Left            =   105
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   9887853
      BorderStyle     =   1
      Appearance      =   1
      MonthBackColor  =   16777215
      ShowToday       =   0   'False
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   69402626
      TitleBackColor  =   8421376
      TitleForeColor  =   16777215
      TrailingForeColor=   12632256
      CurrentDate     =   38456
   End
   Begin VB.Timer fokus 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   2040
   End
End
Attribute VB_Name = "Takvim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Yap As String: Dim GenelMi As Boolean
Private Sub Command1_Click()    'kabul et ve ��k
    Unload Me
End Sub
Private Sub Command2_Click()    'sipari� tarihlerini g�ster
    Write #7, Time$, Me.Name, "Command2_Click", "Start" 'logging
    On Local Error Resume Next
    Me.Enabled = False: Gsipari�.Enabled = False: Me.MousePointer = 11: Gsipari�.List_Gizli.Visible = True
    Gsipari�.List1.Clear
    ''''''''''''''''''''''''''''''
    GenelMi = False
    Yap = "Siparis_Tarihi"
    Arama_islemi
    Gsipari�.List2.Clear: Gsipari�.List3.Clear
    Listeye_Ekle
    ''''''''''''''''''''''''''''''
    Gsipari�.durum.Caption = "G�r�nt�lenen Tarih : " & viev.Month & "." & viev.Year: Gsipari�.List_Gizli.Visible = False
    Me.Enabled = True: Gsipari�.Enabled = True: Me.MousePointer = 1
    If Gsipari�.List1.ListCount >= 1 Then Gsipari�.adet.Caption = Gsipari�.List1.ListCount & " Kay�t Listede G�r�nt�lendi." Else Gsipari�.adet.Caption = "G�r�nt�lenecek Kay�t Yok."
    Write #7, Time$, Me.Name, "Command2_Click", Gsipari�.durum.Caption 'logging
End Sub
Private Sub Command3_Click()    'prova tarihlerini g�ster
    Write #7, Time$, Me.Name, "Command3_Click", "Start" 'logging
    On Local Error Resume Next
    Me.Enabled = False: Gsipari�.Enabled = False: Me.MousePointer = 11: Gsipari�.List_Gizli.Visible = True
    Gsipari�.List1.Clear
    ''''''''''''''''''''''''''''''
    GenelMi = False
    Yap = "Prova_Tarihi"
    Arama_islemi
    Gsipari�.List2.Clear: Gsipari�.List3.Clear
    Listeye_Ekle
    ''''''''''''''''''''''''''''''
    Gsipari�.durum.Caption = "G�r�nt�lenen Tarih : " & viev.Month & "." & viev.Year: Gsipari�.List_Gizli.Visible = False
    Me.Enabled = True: Gsipari�.Enabled = True: Me.MousePointer = 1
    If Gsipari�.List1.ListCount >= 1 Then Gsipari�.adet.Caption = Gsipari�.List1.ListCount & " Kay�t Listede G�r�nt�lendi." Else Gsipari�.adet.Caption = "G�r�nt�lenecek Kay�t Yok."
    Write #7, Time$, Me.Name, "Command3_Click", Gsipari�.durum.Caption 'logging
End Sub
Private Sub Command4_Click()    'teslim tarihlerini g�ster
    Write #7, Time$, Me.Name, "Command4_Click", "Start" 'logging
    On Local Error Resume Next
    Me.Enabled = False: Gsipari�.Enabled = False: Me.MousePointer = 11: Gsipari�.List_Gizli.Visible = True
    Gsipari�.List1.Clear
    ''''''''''''''''''''''''''''''
    GenelMi = False
    Yap = "Teslim_Tarihi"
    Arama_islemi
    Gsipari�.List2.Clear: Gsipari�.List3.Clear
    Listeye_Ekle
    ''''''''''''''''''''''''''''''
    Gsipari�.durum.Caption = "G�r�nt�lenen Tarih : " & viev.Month & "." & viev.Year: Gsipari�.List_Gizli.Visible = False
    Me.Enabled = True: Gsipari�.Enabled = True: Me.MousePointer = 1
    If Gsipari�.List1.ListCount >= 1 Then Gsipari�.adet.Caption = Gsipari�.List1.ListCount & " Kay�t Listede G�r�nt�lendi." Else Gsipari�.adet.Caption = "G�r�nt�lenecek Kay�t Yok."
    Write #7, Time$, Me.Name, "Command4_Click", Gsipari�.durum.Caption 'logging
End Sub
Private Sub Command5_Click()    'yap�lacak t�m i�leri g�ster
    Write #7, Time$, Me.Name, "Command5_Click", "Start" 'logging
    On Local Error Resume Next
    Dim i As Integer
    Me.Enabled = False: Gsipari�.Enabled = False: Me.MousePointer = 11: Gsipari�.List_Gizli.Visible = True
    Gsipari�.List1.Clear
    ''''''''''''''''''''''''''''''
    GenelMi = True: Gsipari�.List4.Clear
    'ilk �nce t�m sipari� tarihlerini liste4e ekler
    Yap = "Siparis_Tarihi"
    Gsipari�.List1.Clear: Gsipari�.List2.Clear: Gsipari�.List3.Clear
    Arama_islemi
    Listeye_Ekle
    'sonra t�m prova tarihlerini liste4e ekler
    Yap = "Prova_Tarihi"
    Gsipari�.List1.Clear: Gsipari�.List2.Clear: Gsipari�.List3.Clear
    Arama_islemi
    Listeye_Ekle
    'en son olarak da t�m teslimat tarihlerini liste4e ekler
    Yap = "Teslim_Tarihi"
    Gsipari�.List1.Clear: Gsipari�.List2.Clear: Gsipari�.List3.Clear
    Arama_islemi
    Listeye_Ekle
    'liste4e eklenmi� i�leri liste1e ekler ve biter
    With Gsipari�
        For i = 0 To .List4.ListCount - 1
            .List1.AddItem .List4.List(i)
        Next i
    End With
    ''''''''''''''''''''''''''''''
    Gsipari�.durum.Caption = "G�r�nt�lenen Tarih : " & viev.Month & "." & viev.Year: Gsipari�.List_Gizli.Visible = False
    Me.Enabled = True: Gsipari�.Enabled = True: Me.MousePointer = 1
    If Gsipari�.List1.ListCount >= 1 Then Gsipari�.adet.Caption = Gsipari�.List1.ListCount & " Kay�t Listede G�r�nt�lendi." Else Gsipari�.adet.Caption = "G�r�nt�lenecek Kay�t Yok."
    Write #7, Time$, Me.Name, "Command5_Click", Gsipari�.durum.Caption 'logging
End Sub

'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Sub Arama_islemi()
    Write #7, Time$, Me.Name, "Arama_islemi", "Start" 'logging
    On Local Error Resume Next
    'bu proced�rde sadece ay�, bizim istedi�imiz ayda olanlar� listeye ekliyoruz
    'daha sonra listeye_ekle proced�r� ile de o kay�tlar i�inden bir arama daha yapaca��z.
    Dim i As Integer: Static j As Byte
    If GenelMi = False Or j = 0 Then For i = 1 To 31: viev.DayBold(i & "." & viev.Month & "." & viev.Year) = False: Next i
    With Anasayfa.dtSipari�.Recordset
        If .RecordCount = 0 Then GoTo son
        .MoveFirst
        For i = 1 To .RecordCount
            If Right(Gsipari�.tarih.Value, 7) = Right(.Fields(Yap), 7) Then
                Gsipari�.List1.AddItem .Fields(Yap)
                viev.DayBold(.Fields(Yap)) = True
            End If
            .MoveNext
        Next i
son:
    End With
    '�u an itibariyle gsipai�.list1 listesine bir y���n tarih oldu.
    Gsipari�.Label2.Caption = ""
    If j < 2 And GenelMi = True Then j = j + 1 Else j = 0
    Write #7, Time$, Me.Name, "Arama_islemi", "End" 'logging
End Sub
Sub Listeye_Ekle()
    Write #7, Time$, Me.Name, "Listeye_Ekle", "Start" 'logging
    On Local Error Resume Next
    Dim i, j As Integer: Dim A_sip As Integer: Dim A_pro As Integer: Dim A_tes As Integer
    With Gsipari�
        .Hangisi.Caption = "1"
        For i = 0 To .List1.ListCount - 1
            A_sip = 0: A_pro = 0: A_tes = 0: .List1.Selected(i) = True
            If i <> 0 Then   'bundan �nce yaz� olmad��� i�in hemen kayda ge�ecek
                For j = 0 To i - 1 'e�er bir �ncekilerde bu yaz� varsa  devam et
                    If .List1.Text = .List1.List(j) Then GoTo Devam 'yani i=i+1
                Next j
            End If
            For j = i To .List1.ListCount - 1   'e�er daha �st sat�rlarda bu yaz� yoksa
                If .List1.Text = .List1.List(j) Then
                    Select Case Yap
                        Case "Siparis_Tarihi": A_sip = A_sip + 1
                        Case "Prova_Tarihi": A_pro = A_pro + 1
                        Case "Teslim_Tarihi": A_tes = A_tes + 1
                    End Select
                End If
            Next j
            Select Case Yap
                Case "Siparis_Tarihi": .List2.AddItem .List1.Text: .List3.AddItem A_sip
                Case "Prova_Tarihi": .List2.AddItem .List1.Text: .List3.AddItem A_pro
                Case "Teslim_Tarihi": .List2.AddItem .List1.Text: .List3.AddItem A_tes
            End Select
Devam:
        Next i
        'liste1 e bir ka� �ey ekledik �imdi de onl�ar� liste2 sayesinde d�zenliyece�iz
        .List1.Clear
        If GenelMi = True Then GoTo Genelse
        For j = 0 To .List2.ListCount - 1
            Select Case Yap
                Case "Siparis_Tarihi": .List1.AddItem .List2.List(j) & " tarihinde  " & .List3.List(j) & " sipari�"
                Case "Prova_Tarihi": .List1.AddItem .List2.List(j) & " tarihinde  " & .List3.List(j) & " prova"
                Case "Teslim_Tarihi": .List1.AddItem .List2.List(j) & " tarihinde  " & .List3.List(j) & " teslimat"
            End Select
        Next j
        Exit Sub
Genelse:
        For j = 0 To .List2.ListCount - 1
            Select Case Yap
                Case "Siparis_Tarihi"
                    .List4.AddItem .List2.List(j) & " tarihinde  " & .List3.List(j) & " sipari�"
                Case "Prova_Tarihi"
                    .List4.AddItem .List2.List(j) & " tarihinde  " & .List3.List(j) & " prova"
                Case "Teslim_Tarihi"
                    .List4.AddItem .List2.List(j) & " tarihinde  " & .List3.List(j) & " teslimat"
            End Select
        Next j
    End With
    Write #7, Time$, Me.Name, "Listeye_Ekle", "End" 'logging
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub viev_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    Gsipari�.tarih.Value = viev.Value
    Gsipari�.G�r_Click
    fokus.Enabled = True
End Sub
Private Sub fokus_Timer()
    Command1.SetFocus
    fokus.Enabled = False
End Sub
Private Sub Form_Load()
    Write #7, Time$, Me.Name, "Form_Load", "Start" 'logging
    Dim Control As Control
    Me.BackColor = rnk_frm_arka
    viev.MonthBackColor = rnk_yaz�_arka: viev.BackColor = rnk_frm_arka: viev.TitleBackColor = rnk_frm_arka: viev.TitleForeColor = rnk_yaz�_�n: viev.TrailingForeColor = rnk_yaz�_�n: viev.ForeColor = rnk_yaz�_�n
    For Each Control In Me
        If TypeOf Control Is OsenXPButton Then Control.BackColor = rnk_btn_arka: Control.BackOver = rnk_btn_arka: Control.ForeColor = rnk_btn_�n: Control.ForeOver = rnk_btn_�n
    Next Control
    Write #7, Time$, Me.Name, "Form_Load", "Succesful" 'logging
End Sub
Private Sub Form_Activate()
    viev.Value = Gsipari�.tarih.Value
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Gsipari�.SetFocus: Gsipari�.takvimon.Caption = ""
    Write #7, Time$, Me.Name, "Form_Unload", "Succesful" 'logging
End Sub


