VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00CDB75F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Malzeme Satisi - Money"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "malzemesatisi.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   Begin MSComCtl2.FlatScrollBar ScrollBar 
      Height          =   3660
      Left            =   11445
      TabIndex        =   23
      Top             =   840
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   6456
      _Version        =   393216
      Appearance      =   0
      LargeChange     =   5
      Orientation     =   1179648
   End
   Begin VB.CommandButton Cmmd7 
      Caption         =   "Çikis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   315
      TabIndex        =   22
      Top             =   4620
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd5 
      Caption         =   "Fiyat Listesi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   315
      TabIndex        =   21
      Top             =   3150
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd4 
      Caption         =   "Stok Kontrolü"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   315
      TabIndex        =   20
      Top             =   2415
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd3 
      Caption         =   "Malzeme Listesi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   315
      TabIndex        =   19
      Top             =   1680
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd2 
      Caption         =   "Malzeme Satisi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   315
      TabIndex        =   18
      Top             =   945
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd6 
      Caption         =   "Hakkinda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   315
      TabIndex        =   17
      Top             =   3885
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd1 
      Caption         =   "Malzeme Girisi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   315
      TabIndex        =   16
      Top             =   210
      Width           =   2985
   End
   Begin VB.CommandButton Command5 
      Caption         =   "listeyi yenile"
      Height          =   540
      Left            =   8715
      TabIndex        =   15
      Top             =   7455
      Width           =   1170
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9460
      TabIndex        =   11
      Top             =   6825
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9460
      TabIndex        =   10
      Top             =   6195
      Width           =   2000
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5670
      TabIndex        =   9
      Top             =   6300
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "sat"
      Height          =   540
      Left            =   7035
      TabIndex        =   8
      Top             =   7455
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ara"
      Height          =   540
      Left            =   5250
      TabIndex        =   7
      Top             =   7455
      Width           =   1170
   End
   Begin VB.ListBox ListAyrinti 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      ItemData        =   "malzemesatisi.frx":030A
      Left            =   3885
      List            =   "malzemesatisi.frx":030C
      TabIndex        =   5
      Top             =   5040
      Width           =   7575
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      ItemData        =   "malzemesatisi.frx":030E
      Left            =   3885
      List            =   "malzemesatisi.frx":0310
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   7575
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   6300
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Saat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8295
      TabIndex        =   14
      Top             =   6825
      Width           =   1170
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarih"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8295
      TabIndex        =   13
      Top             =   6195
      Width           =   1170
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Satis Fiyati"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3885
      TabIndex        =   12
      Top             =   6300
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ayrintilar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3885
      TabIndex        =   6
      Top             =   4620
      Width           =   2955
   End
   Begin VB.Line Line9 
      BorderWidth     =   3
      X1              =   3780
      X2              =   5880
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label YeniMalzemeGirisi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Malzeme Satisi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3780
      TabIndex        =   4
      Top             =   105
      Width           =   2100
   End
   Begin VB.Label LabelKontrol 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   1365
      TabIndex        =   2
      Top             =   6405
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblipucu 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2600
      Left            =   100
      TabIndex        =   1
      Top             =   5750
      Width           =   3500
   End
   Begin VB.Label Labelyazi 
      BackStyle       =   0  'Transparent
      Caption         =   "Bunlari Biliyormuydunuz;"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   200
      TabIndex        =   0
      Top             =   5450
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "malzemesatisi.frx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3750
   End
   Begin VB.Image Background 
      Height          =   495
      Left            =   10560
      Picture         =   "malzemesatisi.frx":268AD
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmmd1_Click()
Form1.Show: Me.Hide
End Sub
Private Sub Cmmd3_Click()
Form3.Show: Me.Hide
End Sub
Private Sub Cmmd4_Click()
Form4.Show: Me.Hide
End Sub
Private Sub Cmmd5_Click()
Form5.Show: Me.Hide
End Sub
Private Sub Cmmd7_Click()
End
End Sub
Private Sub ScrollBar_Change()
On Error Resume Next
ScrollBar.Value = List1.ListIndex
'malzemenin ayrintilarini getirir
AyrintiGetir
List1.SetFocus
End Sub
Private Sub Form_Activate()
'formu sifirliyor ve depodaki malzemeleri listeyi ekler
List1.Clear: ListAyrinti.Clear
Text1 = "": Text2 = "": Text3 = ""
Malzemeler
ZamanBelirt
ScrollBar.Max = List1.ListCount
Cmmd2.SetFocus
End Sub
Private Sub Form_Load()
Background.Left = 3750
Background.Width = Me.Width - 3750
Background.Top = 0
Background.Height = Me.Height
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub List1_Click()
'malzemenin ayrintilarini getirir
AyrintiGetir
ScrollBar.Value = List1.ListIndex
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then ScrollBar.Value = ScrollBar.Value - 1
If KeyCode = 40 Then ScrollBar.Value = ScrollBar.Value + 1
End Sub
Private Sub Timer1_Timer()
'ipucunun deðisme zamani geldi...
Form2.LabelKontrol.Caption = "0"
'ipucunu belirtmek için fonksiyon çaðiriyoruz.
ÝpucuYaz
Timer1.Interval = 4000
End Sub

