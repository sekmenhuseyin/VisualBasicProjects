VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List 
      Height          =   2205
      Index           =   5
      ItemData        =   "Form1.frx":0000
      Left            =   4560
      List            =   "Form1.frx":0002
      TabIndex        =   10
      Top             =   1800
      Width           =   615
   End
   Begin VB.ListBox List 
      Height          =   2205
      Index           =   4
      ItemData        =   "Form1.frx":0004
      Left            =   3840
      List            =   "Form1.frx":0006
      TabIndex        =   9
      Top             =   1800
      Width           =   615
   End
   Begin VB.ListBox List 
      Height          =   2205
      Index           =   3
      ItemData        =   "Form1.frx":0008
      Left            =   3120
      List            =   "Form1.frx":000A
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.ListBox List 
      Height          =   2205
      Index           =   2
      ItemData        =   "Form1.frx":000C
      Left            =   2400
      List            =   "Form1.frx":000E
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.ListBox List 
      Height          =   2205
      Index           =   1
      ItemData        =   "Form1.frx":0010
      Left            =   1680
      List            =   "Form1.frx":0012
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
   Begin VB.ListBox List 
      Height          =   2205
      Index           =   0
      ItemData        =   "Form1.frx":0014
      Left            =   960
      List            =   "Form1.frx":0016
      TabIndex        =   5
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TEMÝZLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ONAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KOLON SAYISI"
      Height          =   195
      Left            =   315
      TabIndex        =   2
      Top             =   1260
      Width           =   1110
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SAYISAL LOTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Byte: Dim j As Byte: Dim k As Byte
Dim diziloto(6) As Byte: Dim Kontrol As Boolean
Private Sub Combo1_Click()
    Command1_Click
End Sub
Private Sub Command1_Click()
    Randomize
    Command2_Click
    ResetLoto
    For i = 1 To Combo1.Text
        For j = 1 To 6
yinele:
            diziloto(j) = Int(Rnd * 49)
            Kontrol = True
            For k = 1 To j - 1
                If diziloto(j) = diziloto(k) Then Kontrol = False
            Next k
            If Kontrol = False Then GoTo yinele
            List(j - 1).AddItem diziloto(j)
            'If MsgBox("çýk?", vbYesNo, "Çýkýþ") = vbYes Then Exit Sub
        Next j
    Next i
    Combo1.SetFocus
End Sub
Private Sub Command2_Click()
    For i = 0 To 5
        List(i).Clear
    Next i
    Combo1.SetFocus
End Sub
Sub ResetLoto()
    For i = 1 To 6
        diziloto(i) = 0
    Next i
End Sub
Private Sub Form_Load()
    For i = 1 To 11
        Combo1.AddItem i
    Next i
    Me.Caption = "Sayýsal Loto Programý"
    Show
    Combo1.ListIndex = 0
    Command1_Click
End Sub

