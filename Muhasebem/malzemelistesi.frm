VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00CDB75F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Malzeme Listesi - Money"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "malzemelistesi.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   Begin VB.CommandButton Cmmd7 
      Caption         =   "��k��"
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
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   3150
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd4 
      Caption         =   "Stok Kontrol�"
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   1680
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd2 
      Caption         =   "Malzeme Sat���"
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
      Top             =   945
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd6 
      Caption         =   "Hakk�nda"
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
      Top             =   3885
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd1 
      Caption         =   "Malzeme Giri�i"
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
      Top             =   210
      Width           =   2985
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   3960
      TabIndex        =   5
      Top             =   6240
      Width           =   7575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4920
      Left            =   3960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   960
      Width           =   7575
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   6300
   End
   Begin VB.Label YeniMalzemeGiri�i 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   360
      Left            =   3780
      TabIndex        =   3
      Top             =   105
      Width           =   2235
   End
   Begin VB.Line Line9 
      BorderWidth     =   3
      X1              =   3795
      X2              =   6050
      Y1              =   510
      Y2              =   510
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
   Begin VB.Label Labelyaz� 
      BackStyle       =   0  'Transparent
      Caption         =   "Bunlar� Biliyormuydunuz;"
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
      Picture         =   "malzemelistesi.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3750
   End
   Begin VB.Image Background 
      Height          =   495
      Left            =   10680
      Picture         =   "malzemelistesi.frx":268A5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
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
Private Sub Cmmd4_Click()
Form4.Show: Me.Hide
End Sub
Private Sub Cmmd5_Click()
Form5.Show: Me.Hide
End Sub
Private Sub Cmmd7_Click()
End
End Sub
Private Sub Form_Activate()
'alt sat�ra atlamay� ��retir
cr$ = Chr$(13) & Chr$(10)
Text1 = ""      'sayfay� temizler
Text2 = ""
Open App.Path + "\stuff\stok.mlz" For Input As #1
Bas:
If EOF(1) Then GoTo Son
'birka� bilgiyi sayfaya yazar
Input #1, mrk, mdl, cns, grn, mkt
Text1.Text = Text1.Text & mrk & "  " & mdl & "    " & cns & "     " & mkt & " Adet" & cr$
GoTo Bas
Son:
Close #1
Cmmd3.SetFocus
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
'ipucunun de�i�me zaman� geldi...
Form3.LabelKontrol.Caption = "0"
'ipucunu belirtmek i�in fonksiyon �a��r�yoruz.
�pucuYaz
Timer1.Interval = 4000
End Sub

