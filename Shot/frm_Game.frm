VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form f_Game 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shot"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9165
   Icon            =   "frm_Game.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fr_Pause 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   6630
      Left            =   9345
      TabIndex        =   1
      Top             =   6825
      Visible         =   0   'False
      Width           =   9150
      Begin MSForms.CommandButton ExitToMenu 
         Height          =   855
         Left            =   2625
         TabIndex        =   8
         Top             =   4410
         Width           =   3795
         Caption         =   "Oyundan Çýk"
         Size            =   "6694;1508"
         FontName        =   "Arial Black"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   162
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton PlayAgain 
         Height          =   855
         Left            =   2670
         TabIndex        =   7
         Top             =   3465
         Width           =   3795
         Caption         =   "Tekrar Oyna"
         Size            =   "6694;1508"
         FontName        =   "Arial Black"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   162
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label lbl_Condition 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "[  ]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4132
         TabIndex        =   3
         Top             =   2340
         Width           =   870
      End
      Begin VB.Label lbl_FinalScore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2265
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   10000
         Y1              =   2835
         Y2              =   2835
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8715
      Top             =   6300
   End
   Begin VB.PictureBox PictureA 
      Height          =   300
      Index           =   0
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   840
      TabIndex        =   5
      Top             =   0
      Width           =   900
   End
   Begin MSForms.Image ImgBall 
      Height          =   480
      Left            =   4620
      Top             =   4935
      Width           =   480
      AutoSize        =   -1  'True
      BorderStyle     =   0
      Size            =   "847;847"
      Picture         =   "frm_Game.frx":137A2
      VariousPropertyBits=   19
   End
   Begin MSForms.CommandButton Raket 
      Height          =   240
      Left            =   3675
      TabIndex        =   6
      Top             =   5670
      Width           =   3165
      Size            =   "5583;423"
      FontHeight      =   165
      FontCharSet     =   162
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Score 
      Height          =   405
      Left            =   1260
      TabIndex        =   4
      Top             =   6240
      Width           =   1200
      BackColor       =   -2147483645
      VariousPropertyBits=   8388627
      Caption         =   "0"
      Size            =   "2117;714"
      FontName        =   "Arial Black"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   162
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label_SC 
      BackStyle       =   0  'Transparent
      Caption         =   "Score :"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   45
      TabIndex        =   0
      Top             =   6240
      Width           =   1200
   End
   Begin VB.Image Ball 
      Height          =   720
      Left            =   2100
      Picture         =   "frm_Game.frx":26F54
      Top             =   -840
      Width           =   720
   End
End
Attribute VB_Name = "f_Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SatýrNo, SütunNo, KutuNo As String
Dim MoveX, MoveY As Integer
Private Sub Timer1_Timer()
    'þimdi topun normal hareketleri
    ImgBall.Left = ImgBall.Left + MoveX
    ImgBall.Top = ImgBall.Top + MoveY
    'hareket alanýna göre gidiþ yönü belirleniyor
    If ImgBall.Left <= 0 Or ImgBall.Left >= ScaleWidth - ImgBall.Width Then MoveX = -MoveX 'formun saðýna ve soluna gitmesini önle
    If ImgBall.Top <= 0 Then MoveY = -MoveY 'formun üstüne gitmesini önle
    If ImgBall.Top >= (Raket.Top - ImgBall.Height) And ImgBall.Left >= (Raket.Left - ImgBall.Width) And ImgBall.Left <= (Raket.Left + Raket.Width) Then MoveY = -MoveY
    If ImgBall.Top >= ScaleHeight Then Call GameOver 'formun altýna kaçarsa oyunu bitir
    'þimdi de top kutucuklara çarptýðýnda iþletilecek kodlar
    Dim i, j, k As Integer: k = -1
    For i = 1 To SütunNo
        For j = 1 To SatýrNo
            k = k + 1
            If PictureA(k).Visible = True And ImgBall.Left >= PictureA(k).Left And ImgBall.Left <= (PictureA(k).Left + PictureA(k).Width) _
            And ImgBall.Top >= PictureA(k).Top And ImgBall.Top <= (PictureA(k).Top + PictureA(k).Height) Then
                PictureA(k).Visible = False
                KutuNo = KutuNo - 1
                MoveY = -MoveY
                Score.Caption = Val(Score.Caption) + 10
            End If
        Next j
    Next i
    'eðer kutular biterse oyunu da bitir
    If KutuNo = 0 Then GameOver
End Sub
Private Sub GameOver()
    Dim i, j As Byte: Dim Best_isim As String
    Timer1.Enabled = False: Call Pause(True)
    For i = 0 To 9
        If Best_Score(i) < Val(Score.Caption) / 10 Then
            Best_isim = InputBox("Bir rekoru kýrdýnýz. Lütfen adýnýzý yazýnýz.")
            If Trim(Best_isim) = "" Then Exit For
            If i <> 9 Then
                For j = 1 To (9 - i)
                    Best_Score(10 - j) = Best_Score(9 - j): Best_Name(10 - j) = Best_Name(9 - j)
                Next j
            End If
            Best_Score(i) = Score.Caption / 10: Best_Name(i) = UpperCaseFirstLetter(Best_isim)
            Exit For
        End If
    Next i
End Sub
Private Sub Pause(Durum As Boolean)
    If Durum = True Then
        If Timer1.Enabled = False Then
            PlayAgain.Caption = "Tekrar Oyna"
            lbl_Condition.Caption = "[ Oyun Bitti ]"
        Else
            PlayAgain.Caption = "Devam Et"
            lbl_Condition.Caption = "[ Oyun Duraklatýldý ]"
        End If
        lbl_FinalScore.Caption = "Score :  " & Score.Caption: Fr_Pause.Visible = True: Fr_Pause.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        lbl_Condition.Left = (Me.Width - lbl_Condition.Width) / 2: Timer1.Enabled = False: PlayAgain.SetFocus
    Else
        Fr_Pause.Visible = False: Fr_Pause.Move ScaleWidth, ScaleHeight: Timer1.Enabled = True
    End If
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub Raket_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Raket.Left = X
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Raket.Left = X - Raket.Width / 2
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyLeft: If Raket.Left >= 15 Then Raket.Left = Raket.Left - 15 Else Raket.Left = 0
        Case vbKeyRight: If Raket.Left >= (Me.ScaleWidth - Raket.Width - 15) Then Raket.Left = Raket.Left + 15 Else Raket.Left = Me.ScaleWidth - Raket.Width
        Case vbKeyEscape: If Fr_Pause.Visible = False Then Pause (True) Else Unload Me
    End Select
End Sub
Private Sub Form_Load()
    Dim i, j, k As Integer: Randomize
    SatýrNo = 5: SütunNo = 10: KutuNo = Val(SatýrNo) * Val(SütunNo): PictureA(0).BackColor = QBColor(10)
    Width = PictureA(0).Width * SütunNo + 30: k = 0
    For i = 1 To SatýrNo
        For j = 1 To SütunNo
            k = k + 1: Load PictureA(k)
            PictureA(k).Left = PictureA(k - 1).Left + PictureA(k).Width
            PictureA(k).Top = (i - 1) * PictureA(k).Height
            PictureA(k).BackColor = QBColor(j)
            PictureA(k).Visible = True
        Next j
        PictureA(k).Left = 0: PictureA(k).Top = (i) * PictureA(k).Height
    Next i
    Unload PictureA(k): k = k - 1
    Fr_Pause.Visible = False: Fr_Pause.Move ScaleWidth, ScaleHeight: Score.Caption = "0"
    ImgBall.Top = Int(Rnd * (Raket.Top - PictureA(k).Top)) + PictureA(k).Top: ImgBall.Left = Int(Rnd * (ScaleWidth - 15))
    MoveX = -50: MoveY = -50: Timer1.Enabled = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    f_Main.Top = Me.Top: f_Main.Left = Me.Left: f_Main.Show
End Sub
Private Sub PlayAgain_Click()
    On Error Resume Next
    Select Case PlayAgain.Caption
        Case "Tekrar Oyna"
            Dim i, j, k As Integer
            For i = 1 To SatýrNo
                For j = 1 To SütunNo
                    k = k + 1: Unload PictureA(k)
                Next j
            Next i
            Form_Load
        Case "Devam Et": Pause (False)
    End Select
End Sub
Private Sub ExitToMenu_Click()
    Unload Me
End Sub
