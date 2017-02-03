VERSION 5.00
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form Ödeme 
   BackColor       =   &H0096E06D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ödeme Zamanlarý"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
   Icon            =   "frmÖdeme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H0096E06D&
      Caption         =   "Ödemeler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   2835
      TabIndex        =   8
      Top             =   630
      Width           =   3400
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   2175
         ItemData        =   "frmÖdeme.frx":0442
         Left            =   180
         List            =   "frmÖdeme.frx":0444
         TabIndex        =   10
         Top             =   555
         Width           =   1725
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   2175
         ItemData        =   "frmÖdeme.frx":0446
         Left            =   1920
         List            =   "frmÖdeme.frx":0448
         TabIndex        =   9
         Top             =   555
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ödeme Miktarý"
         Height          =   195
         Left            =   1920
         TabIndex        =   12
         Top             =   345
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tarih"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   210
      Top             =   4410
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0096E06D&
      Caption         =   "Genel Bilgi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   120
      TabIndex        =   1
      Top             =   615
      Width           =   2655
      Begin VB.TextBox kalan 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   "0"
         Top             =   705
         Width           =   1695
      End
      Begin VB.TextBox fiyat 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Text            =   "0"
         Top             =   345
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Borcu"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   750
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tutar"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   390
         Width           =   375
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   -1  'True
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   -1  'True
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   -1  'True
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin OsenXPCntrl.OsenXPButton Command1 
      Cancel          =   -1  'True
      Height          =   840
      Left            =   210
      TabIndex        =   6
      Top             =   2625
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
      MICON           =   "frmÖdeme.frx":044A
      PICN            =   "frmÖdeme.frx":0466
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Müþterinin Adý Soyadý"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1155
      TabIndex        =   7
      Top             =   135
      Width           =   1905
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   120
      X2              =   6240
      Y1              =   500
      Y2              =   500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ad Soyad :"
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   195
      Width           =   780
   End
End
Attribute VB_Name = "Ödeme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fiyat_GotFocus(): Call SelectAllText: End Sub
Private Sub kalan_GotFocus(): Call SelectAllText: End Sub
Private Sub List1_Click()
    List2.ListIndex = List1.ListIndex
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    List2.ListIndex = List1.ListIndex
End Sub
Private Sub List2_Click()
    List1.ListIndex = List2.ListIndex
End Sub
Private Sub List2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    List1.ListIndex = List2.ListIndex
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Write #7, Time$, Me.Name, "Form_Load", "Start" 'logging
    Dim Control As Control: Me.BackColor = rnk_frm_arka
    For Each Control In Me
        If TypeOf Control Is OsenXPButton Then Control.BackColor = rnk_btn_arka: Control.BackOver = rnk_btn_arka: Control.ForeColor = rnk_btn_ön: Control.ForeOver = rnk_btn_ön
        If TypeOf Control Is Label Then Control.BackColor = rnk_frm_arka: Control.ForeColor = rnk_frm_ön
        If TypeOf Control Is Frame Then Control.BackColor = rnk_frm_arka: Control.ForeColor = rnk_frm_ön
        If TypeOf Control Is TextBox Then Control.BackColor = rnk_yazý_arka: Control.ForeColor = rnk_yazý_ön: Control.Locked = True
        If TypeOf Control Is ListBox Then Control.BackColor = rnk_yazý_arka: Control.ForeColor = rnk_yazý_ön
        If TypeOf Control Is Line Then Control.BorderColor = rnk_frm_ön
    Next Control
    Write #7, Time$, Me.Name, "Form_Load", "Successful" 'logging
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Göster.Enabled = True
    Write #7, Time$, Me.Name, "Form_Unload", "Successful" 'logging
End Sub
Private Sub Timer1_Timer()
    Write #7, Time$, Me.Name, "Timer1_Timer", "Start" 'logging
    On Local Error Resume Next
    Dim i As Integer: Dim M_ID, S_ID As String
    If Dir(App.path + "\temp2.txt") <> "" Then
        Open App.path + "\temp2.txt" For Input As #2: Input #2, M_ID: Input #2, S_ID: Close #2
    End If
    Label6.Caption = "  " + Anasayfa.dtMüþteri.Recordset.Fields("Musteri_Adi") + "  " + Anasayfa.dtMüþteri.Recordset.Fields("Musteri_Soyadi") + "  "
    With Anasayfa.dtSipariþ.Recordset
        .MoveFirst
        For i = 1 To .RecordCount
            If Val(.Fields("Musteri_Kodu")) = Val(M_ID) Then fiyat.Text = Val(fiyat.Text) + Val(.Fields("Ucret"))
            .MoveNext
        Next i
    End With: DoEvents
    With Anasayfa.dtÖdeme.Recordset
        .MoveFirst
        For i = 1 To .RecordCount
            If .Fields("Musteri_Kodu") = M_ID Then
                kalan.Text = Val(kalan.Text) + Val(.Fields("Odenen_Fiyat"))
                List1.AddItem .Fields("Tarih")
                List2.AddItem .Fields("Odenen_Fiyat")
            End If
            .MoveNext
        Next i
    End With: DoEvents
    kalan.Text = Val(fiyat.Text) - Val(kalan.Text)
    With Anasayfa.dtSipariþ.Recordset 'sipariþler
        .MoveFirst: While CStr(.Fields("Siparis_Kodu")) <> S_ID: .MoveNext: Wend
    End With
    Write #7, Time$, Me.Name, "Timer1_Timer", "Müþteri:" & M_ID 'logging
    Timer1.Enabled = False
End Sub


