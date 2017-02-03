VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00CDB75F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fiyat Listesi - Money"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15525
   Icon            =   "fiyatlistesi.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   15525
   Begin VB.CommandButton Cmmd7 
      Caption         =   "Çýkýþ"
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
      TabIndex        =   9
      Top             =   4590
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   1680
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd2 
      Caption         =   "Malzeme Satýþý"
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
      TabIndex        =   5
      Top             =   945
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd6 
      Caption         =   "Hakkýnda"
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
      TabIndex        =   4
      Top             =   3885
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd1 
      Caption         =   "Malzeme Giriþi"
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
      TabIndex        =   3
      Top             =   210
      Width           =   2985
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   6300
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
   Begin VB.Label Labelyazý 
      BackStyle       =   0  'Transparent
      Caption         =   "Bunlarý Biliyormuydunuz;"
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
      Picture         =   "fiyatlistesi.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3750
   End
   Begin VB.Image Background 
      Height          =   495
      Left            =   10800
      Picture         =   "fiyatlistesi.frx":268A5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmmd1_Click()
Form1.Show: Me.Hide
End Sub
Private Sub Cmmd2_Click()
Form2.Show: Me.Hide
End Sub
Private Sub Cmmd3_Click()
Form3.Show: Me.Hide
End Sub
Private Sub Cmmd4_Click()
Form4.Show: Me.Hide
End Sub
Private Sub Cmmd7_Click()
End
End Sub
Private Sub Form_Activate()
Cmmd5.SetFocus
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
Private Sub Timer1_Timer()
'ipucunun deðiþme zamaný geldi...
Form5.LabelKontrol.Caption = "0"
'ipucunu belirtmek için fonksiyon çaðýrýyoruz.
ÝpucuYaz
Timer1.Interval = 4000
End Sub

