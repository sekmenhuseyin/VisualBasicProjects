VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9165
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":1562A
   ScaleHeight     =   5655
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   420
      Top             =   1050
   End
   Begin MCI.MMControl mp3 
      Height          =   435
      Left            =   315
      TabIndex        =   3
      Top             =   1785
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   767
      _Version        =   393216
      RecordMode      =   0
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   1050
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1680
      Top             =   210
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1260
      Top             =   210
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   840
      Top             =   210
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   420
      Top             =   210
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   210
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   0
      Pattern         =   "*.dsc"
      TabIndex        =   1
      Top             =   0
      Width           =   6105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   1470
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   0
      TabIndex        =   2
      Top             =   1785
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   840
      TabIndex        =   0
      Top             =   1785
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Zaman�M�, k�rm�z�, ye�il, mavi, ka�defaD�nd� As Integer
Dim Neye, renk, nereye, mp3iSim As String
Dim san, taym�r, taym4, a As Integer
Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        File1_MouseDown 1, Shift, 0, 0
    Case 27
        End
    Case 32
        File1_MouseDown 1, Shift, 0, 0
    End Select
End Sub
Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Open App.Path + "\mp3\" + File1.List(File1.ListIndex) For Input As #1
        Input #1, mp3iSim, san
        If Dir(App.Path + "\mp3\" + mp3iSim) = "" Then MsgBox "Bilgi dosyas� bozuk !", vbExclamation: Close #1: Exit Sub
        File1.Visible = False: Me.BackColor = san
        Input #1, taym�r, taym4, san
    k�rm�z� = 0: mavi = 0: ye�il = 0: Zaman�M� = 0: a = 0: Neye = "Beyaza": ka�defaD�nd� = 0: Timer3.Interval = 50: nereye = "k�rm�z�ya"
    mp3.FileName = App.Path + "\mp3\" + mp3iSim: mp3.Command = "Open": Me.MousePointer = 99
    Timer7.Enabled = True
End Sub
Private Sub Form_Click()
    If Timer6.Enabled = True Or Timer7.Enabled = True Or Timer3.Enabled = True Then
        nereye = "Beyaza": Timer1.Enabled = False: Timer2.Enabled = False: Timer3.Enabled = True: Timer5.Enabled = False: Timer6.Enabled = False
    Else
        End
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 27
        Form_Click
    End Select
End Sub
Private Sub Form_Load()
    Zaman�M� = 0: Neye = "Beyaza": nereye = "k�rm�z�ya": ka�defaD�nd� = 0: k�rm�z� = 0: mavi = 0: ye�il = 0: renk = "&H000000FF&": a = 0
    Me.Move 0, 0, Screen.Width, Screen.Height
    File1.Width = Me.Width: File1.Path = App.Path + "\mp3\"
    If File1.ListCount = 0 Then File1.Height = 226 Else File1.Height = Val(File1.ListCount) * 226: File1.Selected(0) = True
    If File1.Height > Me.Height / 2 Then File1.Height = Me.Height / 2
End Sub
Private Sub Label1_Change()
    Label2.Caption = Val(Label2.Caption) + 1
End Sub
Private Sub Label3_Change()
    a = a + 1
    If a = 3 Then Timer6.Enabled = True: mp3.Command = "play": Timer7.Enabled = False: Label1.Caption = Time: Label2.Caption = 0: a = 0
End Sub
Private Sub Timer1_Timer() 'ani siyah beyaz de�i�imi
    Select Case Zaman�M�
    Case 12: Me.BackColor = &HFFFFFF: Zaman�M� = 0: k�rm�z� = 15: ye�il = 15: mavi = 15
    Case Else: Me.BackColor = 0: Zaman�M� = Zaman�M� + 1: k�rm�z� = 0: ye�il = 0: mavi = 0
    End Select
End Sub
Private Sub Timer2_Timer() 'siyaha ve beyaza yava� ge�i�
    If a < 4 And Me.BackColor = 0 Then
        a = a + 1
        Exit Sub
    ElseIf a = 4 And Me.BackColor = 0 Then
        a = 0
    End If
    Select Case Neye
    Case "Siyaha"
        If mavi <> 0 Then mavi = mavi - 1
        If k�rm�z� <> 0 Then k�rm�z� = k�rm�z� - 1
        If ye�il <> 0 Then ye�il = ye�il - 1
        If k�rm�z� = 0 And ye�il = 0 And mavi = 0 Then Neye = "Beyaza"
    Case "Beyaza"
        If k�rm�z� <> 15 Then k�rm�z� = k�rm�z� + 1
        If ye�il <> 15 Then ye�il = ye�il + 1
        If mavi <> 15 Then mavi = mavi + 1
        If k�rm�z� = 15 And ye�il = 15 And mavi = 15 Then Neye = "Siyaha"
    End Select
    Me.BackColor = RGB(k�rm�z� * 11, ye�il * 11, mavi * 11)
End Sub
Private Sub Timer3_Timer() 'renkler aras� yava� ge�i�
    Select Case nereye
    Case "k�rm�z�ya"
        If k�rm�z� <> 15 Then k�rm�z� = k�rm�z� + 1
        If mavi <> 0 Then mavi = mavi - 1
        If k�rm�z� = 15 Then nereye = "ye�ile"
    Case "ye�ile"
        If ye�il <> 15 Then ye�il = ye�il + 1
        If k�rm�z� <> 0 Then k�rm�z� = k�rm�z� - 1
        If ye�il = 15 Then nereye = "maviye"
    Case "maviye"
        If mavi <> 15 Then mavi = mavi + 1
        If ye�il <> 0 Then ye�il = ye�il - 1
        If mavi = 15 Then nereye = "k�rm�z�ya"
    Case "Beyaza"
        If mavi <> 15 Then mavi = mavi + 1
        If ye�il <> 15 Then ye�il = ye�il + 1
        If k�rm�z� <> 15 Then k�rm�z� = k�rm�z� + 1
        If k�rm�z� = 15 And mavi = 15 And ye�il = 15 Then BirM�zikSonu: Exit Sub
    End Select
    Me.BackColor = RGB(k�rm�z� * 11, ye�il * 11, mavi * 11)
End Sub
Private Sub Timer4_Timer() 'aras�ra fonu siyah yapar
    Me.BackColor = 0
End Sub
Private Sub Timer5_Timer() 'beyaz-renkli kar���m� ge�i�siz
    Select Case Zaman�M�
    Case 10: Me.BackColor = &HFFFFFF: Zaman�M� = 0: ka�defaD�nd� = ka�defaD�nd� + 1
    Case 9:  Me.BackColor = &HFFFFFF: Zaman�M� = Zaman�M� + 1
    Case 8:  Me.BackColor = &HFFFFFF: Zaman�M� = Zaman�M� + 1
    Case 7:  Me.BackColor = &HFFFFFF: Zaman�M� = Zaman�M� + 1
    Case 6:  Me.BackColor = &HFFFFFF: Zaman�M� = Zaman�M� + 1
    Case Else: Me.BackColor = Val(renk): Zaman�M� = Zaman�M� + 1
    End Select
    If ka�defaD�nd� = 7 Then
        renk = "&H0000FF00&": k�rm�z� = 0: ye�il = 0: mavi = 15
    ElseIf ka�defaD�nd� = 14 Then
        renk = "&H00FF0000&": k�rm�z� = 0: ye�il = 15: mavi = 0
    ElseIf ka�defaD�nd� >= 21 Then
        renk = "&H000000FF&": k�rm�z� = 15: ye�il = 0: mavi = 0
        ka�defaD�nd� = 0
    End If
End Sub
Private Sub Timer6_Timer()
    Select Case taym�r
    Case 0
        Timer3.Interval = 100: nereye = "Beyaza": Timer1.Enabled = False: Timer2.Enabled = False: Timer3.Enabled = True: Timer5.Enabled = False: Timer6.Enabled = False
    Case 1
        Timer1.Enabled = True: Timer2.Enabled = False: Timer3.Enabled = False: Timer5.Enabled = False
        a = 0: Neye = "Beyaza": k�rm�z� = 0: mavi = 0: ye�il = 0: ka�defaD�nd� = 0: Timer3.Interval = 50
    Case 2
        Timer1.Enabled = False: Timer2.Enabled = True: Timer3.Enabled = False: Timer5.Enabled = False
        Zaman�M� = 0: ka�defaD�nd� = 0: Timer3.Interval = 50
    Case 3
        Timer1.Enabled = False: Timer2.Enabled = False: Timer3.Enabled = True: Timer5.Enabled = False
        Zaman�M� = 0: a = 0: Neye = "Beyaza": ka�defaD�nd� = 0
    Case 5
        Timer1.Enabled = False: Timer2.Enabled = False: Timer3.Enabled = False: Timer5.Enabled = True
        a = 0: Neye = "Beyaza": k�rm�z� = 0: mavi = 0: ye�il = 0: Timer3.Interval = 50
    End Select
    If taym4 = 4 Then Timer4.Enabled = True Else Timer4.Enabled = False
    If Val(Label2.Caption) = Val(san) Then
        Input #1, taym�r, taym4, san
    End If
    Label1.Caption = Time
End Sub
Sub BirM�zikSonu()
    Timer1.Enabled = False: Timer2.Enabled = False: Timer3.Enabled = False: Timer5.Enabled = False: Timer6.Enabled = False: Timer7.Enabled = False
    mp3.Command = "Close": Close #1: Me.MousePointer = 0: Me.BackColor = &HFFFFFF: File1.Visible = True
End Sub
Private Sub Timer7_Timer()
    Label3.Caption = Time
End Sub
Private Sub mp3_Done(NotifyCode As Integer)
    If NotifyCode = 1 Then Form_Click
End Sub

