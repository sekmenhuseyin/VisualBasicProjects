VERSION 5.00
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form Anasayfa 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   9885
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7140
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmHome.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   5985
      Top             =   6825
   End
   Begin OsenXPCntrl.OsenXPButton Command1 
      Default         =   -1  'True
      Height          =   975
      Left            =   255
      TabIndex        =   0
      Top             =   1425
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Sipariþ Oluþtur"
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
      MICON           =   "frmHome.frx":000C
      PICN            =   "frmHome.frx":0028
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
      Height          =   975
      Left            =   255
      TabIndex        =   1
      Top             =   2505
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Arama"
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
      MICON           =   "frmHome.frx":0902
      PICN            =   "frmHome.frx":091E
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
      Height          =   975
      Left            =   255
      TabIndex        =   2
      Top             =   3600
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Günlük Sipariþ Listesi"
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
      MICON           =   "frmHome.frx":0D70
      PICN            =   "frmHome.frx":0D8C
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
      Height          =   975
      Left            =   255
      TabIndex        =   3
      Top             =   4665
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Mali Ýþlemler"
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
      MICON           =   "frmHome.frx":1666
      PICN            =   "frmHome.frx":1682
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
      Height          =   975
      Left            =   255
      TabIndex        =   4
      Top             =   5745
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Ayarlar"
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
      MICON           =   "frmHome.frx":1AD4
      PICN            =   "frmHome.frx":1AF0
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
      Height          =   975
      Left            =   255
      TabIndex        =   5
      Top             =   6825
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Hakkýnda"
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
      MICON           =   "frmHome.frx":16C62
      PICN            =   "frmHome.frx":16C7E
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
      Cancel          =   -1  'True
      Height          =   975
      Left            =   255
      TabIndex        =   6
      Top             =   7905
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Çýkýþ"
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
      MICON           =   "frmHome.frx":170D0
      PICN            =   "frmHome.frx":170EC
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Data dtSipariþ 
      Caption         =   "Sipariþler"
      Connect         =   "Access"
      DatabaseName    =   "G:\Visual Basic\Denemelerim\EsraM\include\data\Data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4305
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tbl_Siparisler"
      Top             =   8820
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data dtSipariþTürü 
      Caption         =   "Sipariþ Türleri"
      Connect         =   "Access"
      DatabaseName    =   "G:\Visual Basic\Denemelerim\EsraM\include\data\Data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4305
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tbl_Siparis_Turleri"
      Top             =   7245
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data dtResim 
      Caption         =   "Resimler"
      Connect         =   "Access"
      DatabaseName    =   "G:\Visual Basic\Denemelerim\EsraM\include\data\Data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4305
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tbl_Resimler"
      Top             =   7560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data dtÖdeme 
      Caption         =   "Ödemeler"
      Connect         =   "Access"
      DatabaseName    =   "G:\Visual Basic\Denemelerim\EsraM\include\data\Data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4305
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tbl_Odemeler"
      Top             =   7875
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data dtMüþteri 
      Caption         =   "Müþteriler"
      Connect         =   "Access"
      DatabaseName    =   "G:\Visual Basic\Denemelerim\EsraM\include\data\Data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4305
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tbl_Musteriler"
      Top             =   8190
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data dtKumaþTürü 
      Caption         =   "Kumaþ Türleri"
      Connect         =   "Access"
      DatabaseName    =   "G:\Visual Basic\Denemelerim\EsraM\include\data\Data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4305
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tbl_KumasTurleri"
      Top             =   8505
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label_Genel 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   15
      TabIndex        =   8
      Top             =   15
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   1110
      Left            =   120
      Picture         =   "frmHome.frx":179C6
      Top             =   90
      Width           =   3555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   555
      Left            =   4080
      TabIndex        =   7
      Top             =   345
      Width           =   2145
   End
   Begin VB.Image Image2 
      Height          =   9240
      Left            =   0
      Picture         =   "frmHome.frx":18045
      Top             =   0
      Width           =   6600
   End
   Begin VB.Menu jmenu 
      Caption         =   "Pop Up"
      Visible         =   0   'False
      Begin VB.Menu altcommand1 
         Caption         =   "Sipariþ Oluþtur"
         Shortcut        =   {F2}
      End
      Begin VB.Menu altcommand2 
         Caption         =   "Arama"
         Shortcut        =   {F3}
      End
      Begin VB.Menu altcommand3 
         Caption         =   "Günlük Sipariþ Listesi"
         Shortcut        =   {F4}
      End
      Begin VB.Menu altcommand4 
         Caption         =   "Mali Ýþlemler"
         Shortcut        =   {F5}
      End
      Begin VB.Menu alttire0 
         Caption         =   "-"
      End
      Begin VB.Menu altcommand5 
         Caption         =   "Ayarlar"
         Shortcut        =   {F1}
      End
      Begin VB.Menu alttire1 
         Caption         =   "-"
      End
      Begin VB.Menu altTheme 
         Caption         =   "Tema"
         Begin VB.Menu altTema 
            Caption         =   "EsraM Standart"
            Index           =   0
         End
      End
      Begin VB.Menu alttire2 
         Caption         =   "-"
      End
      Begin VB.Menu altExit 
         Caption         =   "Çýkýþ"
      End
   End
End
Attribute VB_Name = "Anasayfa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub Form_Activate(): On Error Resume Next: Call frmMain.MDIForm_Resize: End Sub
Public Sub AnaGörünüm()
    Write #7, Time$, Me.Name, "AnaGörünüm", "Start" 'logging
    Dim Control As Control: Dim i As Integer
    'arkaplan resmi ve logo
    If Dir(Tema_Yeri & "\logo.gif") <> "" Then Image1.Picture = LoadPicture(Tema_Yeri & "\logo.gif")
    If Dir(Tema_Yeri & "\Background.jpg") <> "" Then Image2.Picture = LoadPicture(Tema_Yeri & "\Background.jpg")
    Me.BackColor = rnk_frm_arka: frmMain.BackColor = rnk_btn_arka: Anasayfa.BackColor = rnk_btn_arka
    For Each Control In Me
        If TypeOf Control Is OsenXPButton Then Control.BackColor = rnk_btn_arka: Control.BackOver = rnk_btn_arka: Control.ForeColor = rnk_btn_ön: Control.ForeOver = rnk_btn_ön
        If TypeOf Control Is Label Then Control.ForeColor = rnk_frm_ön
        If TypeOf Control Is Line Then Control.BorderColor = rnk_frm_ön
        If TypeOf Control Is Data Then Control.DatabaseName = App.path + "\include\data\Data.mdb"
    Next Control
    For i = 0 To altTema.Count - 1
        If Tema_Adý = Themes(i).TemaAd Then altTema(i).Checked = True Else altTema(i).Checked = False
    Next i
    Write #7, Time$, Me.Name, "AnaGörünüm", "End" 'logging
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27: Call Command3_Click 'exit [esc]
        Case 112: Call Command6_Click 'ayarlar [f1]
        Case 113: Call Command1_Click 'sipariþ [f2]
        Case 114: Call Command4_Click 'arama [f3]
        Case 115: Call Command2_Click 'gsipariþ [f4]
        Case 116: Command7_Click 'maliiþlemler [f5]
    End Select
End Sub
Private Sub Form_Load()
    Write #7, Time$, Me.Name, "Form_Load", "Start" 'logging
    Me.Move 0, 0, Image2.Width, Image2.Height: Call AnaGörünüm: Label_Genel.Move 0, 0, ScaleWidth, ScaleHeight
    Write #7, Time$, Me.Name, "Form_Load", "Successful" 'logging
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+

Private Sub Timer1_Timer()
    Dim Control As Control
    For Each Control In Me
        If TypeOf Control Is Data Then
            Do While Control.Recordset.EOF <> True: Control.Recordset.MoveNext: Loop
            If Control.Recordset.BOF <> True Then Control.Recordset.MoveFirst
        End If
    Next Control
    Timer1.Enabled = False
End Sub
Private Sub altTema_Click(Index As Integer)
    Tema_Yeri = Themes(Index).TemaDizin: Tema_Adý = Themes(Index).TemaAd
    Write #7, Time$, Me.Name, "altTema_Click", Tema_Adý 'logging
    rnk_yazý_arka = ReadStringFromIni("Theme", "rnk_yazý_arka", CStr(rnk_yazý_arka), Tema_Yeri)
    rnk_frm_arka = ReadStringFromIni("Theme", "rnk_frm_arka", CStr(rnk_frm_arka), Tema_Yeri)
    rnk_yazý_ön = ReadStringFromIni("Theme", "rnk_yazý_ön", CStr(rnk_yazý_ön), Tema_Yeri)
    rnk_btn_arka = ReadStringFromIni("Theme", "rnk_btn_arka", CStr(rnk_btn_arka), Tema_Yeri)
    rnk_btn_ön = ReadStringFromIni("Theme", "rnk_btn_ön", CStr(rnk_btn_ön), Tema_Yeri)
    Call AnaGörünüm: cmbindex = Index + 1
    Write #7, Time$, Me.Name, "altTema_Click", "End" 'logging
End Sub
Private Sub altcommand1_Click(): Command1_Click: End Sub
Private Sub altcommand2_Click(): Command4_Click: End Sub
Private Sub altcommand3_Click(): Command2_Click: End Sub
Private Sub altcommand4_Click(): Command7_Click: End Sub
Private Sub altcommand5_Click(): Command6_Click: End Sub
Private Sub altExit_Click(): Command3_Click: End Sub
Private Sub Command1_Click(): Anasayfa.Visible = False: Sipariþ.Show: End Sub
Private Sub Command4_Click(): Anasayfa.Visible = False: Arama.Show: End Sub
Private Sub Command2_Click(): Anasayfa.Visible = False: Gsipariþ.Show: End Sub
Private Sub Command7_Click(): Anasayfa.Visible = False: Maliiþlemler.Show: End Sub
Private Sub Command6_Click(): Me.Enabled = False: Ayarlar.Show: End Sub
Private Sub Command5_Click(): Me.Enabled = False: Hakkýnda.Show: End Sub
Private Sub Command3_Click(): The_End: End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*'popupmenu
Private Sub Label_Genel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    jmenu.Visible = False: If Button = 2 Then Exit Sub
End Sub
Private Sub Label_Genel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If jmenu.Visible = False And Button = 2 Then PopupMenu jmenu
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*'popupmenu
