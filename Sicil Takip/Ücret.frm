VERSION 5.00
Begin VB.Form Ücret 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ücret Deðiþikliði"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   Icon            =   "Ücret.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Kaydet"
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
      Left            =   6825
      TabIndex        =   5
      Top             =   4095
      Width           =   2535
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
      Left            =   420
      TabIndex        =   0
      Top             =   4095
      Width           =   2535
   End
   Begin VB.TextBox Text1 
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
      IMEMode         =   3  'DISABLE
      Left            =   4935
      MaxLength       =   50
      TabIndex        =   1
      Top             =   630
      Width           =   4500
   End
   Begin VB.TextBox Text4 
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
      Left            =   4935
      TabIndex        =   4
      Top             =   3255
      Width           =   4500
   End
   Begin VB.TextBox Text3 
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
      Left            =   4935
      TabIndex        =   3
      Top             =   2310
      Width           =   4500
   End
   Begin VB.TextBox Text2 
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
      Left            =   4935
      TabIndex        =   2
      Top             =   1470
      Width           =   4500
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hastalýk Sigortasý Prim Yüzdesi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   405
      TabIndex        =   9
      Top             =   3255
      Width           =   4530
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Kaza Sigortasý Prim Yüzdesi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   405
      TabIndex        =   8
      Top             =   2310
      Width           =   4530
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "16 Yaþ Üstü Günlük Ücret"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   405
      TabIndex        =   7
      Top             =   1470
      Width           =   4530
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "16 Yaþ Altý Günlük Ücret"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   405
      TabIndex        =   6
      Top             =   630
      Width           =   4530
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Height          =   4600
      Left            =   200
      TabIndex        =   10
      Top             =   200
      Width           =   9600
   End
End
Attribute VB_Name = "Ücret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Giriþ.Show: Unload Me
End Sub
Private Sub Command2_Click()
If Text1 = "" Then MsgBox "Eksik Bilgi Girdiniz": Text1.SetFocus: GoTo son
If Text2 = "" Then MsgBox "Eksik Bilgi Girdiniz": Text2.SetFocus: GoTo son
If Text3 = "" Then MsgBox "Eksik Bilgi Girdiniz": Text3.SetFocus: GoTo son
If Text4 = "" Then MsgBox "Eksik Bilgi Girdiniz": Text4.SetFocus: GoTo son
Open App.Path + "\Salary.sys" For Output As #1
Write #1, Text1, Text2, Text3, Text4
MsgBox "Bilgiler Kaydedildi."
Close #1
son:
Command1.SetFocus
End Sub
Private Sub Form_Activate()
Open App.Path + "\Salary.sys" For Input As #1
If EOF(1) Then GoTo son
Input #1, a, b, c, d
Text1 = a: Text2 = b: Text3 = c: Text4 = d
son:
Close #1
Text1.Text = Format(Text1, "###,###,###,###")
Text2.Text = Format(Text2, "###,###,###,###")
End Sub
Private Sub Form_Unload(Cancel As Integer)
Command1_Click
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text1.Text = Format(Text1, "###,###,###,###"): Text2.SetFocus
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.Text = Format(Text2, "###,###,###,###"): Text3.SetFocus
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text4.SetFocus
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command2_Click
End Sub

