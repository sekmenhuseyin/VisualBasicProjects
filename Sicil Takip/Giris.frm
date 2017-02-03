VERSION 5.00
Begin VB.Form Giris 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5010
   ClientLeft      =   1350
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "Giris.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Ýslem Sonu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   1343
      Picture         =   "Giris.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3780
      Width           =   5625
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Kimlik Listesi"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   1343
      Picture         =   "Giris.frx":2BE4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2930
      Width           =   5625
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ücret Deðisikliði"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   1343
      Picture         =   "Giris.frx":5386
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2080
      Width           =   5625
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Yazýcýya Gönder"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   1343
      Picture         =   "Giris.frx":57EC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1230
      Width           =   5625
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kimlik Ýslemleri"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   1343
      Picture         =   "Giris.frx":5C2E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   380
      Width           =   5625
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Height          =   4635
      Left            =   570
      TabIndex        =   5
      Top             =   188
      Width           =   7170
   End
End
Attribute VB_Name = "Giris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Kayýtlar.Show: Me.Hide
End Sub
Private Sub Command2_Click()
Yazdýr.Show: Me.Hide
End Sub
Private Sub Command3_Click()
Ücret.Show: Me.Hide
End Sub
Private Sub Command4_Click()
Kimlik.Show: Me.Hide
End Sub
Private Sub Command5_Click()
End
End Sub
Private Sub Form_Load()
bosluksayisi = Int((Me.Width - 4238) / 60)
Me.Caption = "100. YIL EML." & String$(bosluksayisi, " ") & "3308 ÝÞLEMLERÝ"
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
