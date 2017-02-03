VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form AntiDsc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anti Disco"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   Icon            =   "AntiDsc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4740
      ItemData        =   "AntiDsc.frx":1562A
      Left            =   3150
      List            =   "AntiDsc.frx":1562C
      TabIndex        =   9
      Top             =   1050
      Width           =   4320
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Dosyaya Kaydet"
      Enabled         =   0   'False
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
      Left            =   105
      TabIndex        =   8
      Top             =   3675
      Width           =   2955
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Zamaný Ekle"
      Enabled         =   0   'False
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
      Left            =   105
      TabIndex        =   6
      Top             =   1575
      Width           =   2850
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Baþla"
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
      Left            =   105
      TabIndex        =   5
      Top             =   1050
      Width           =   2850
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   0
      Pattern         =   "*.mp3"
      TabIndex        =   1
      Top             =   0
      Width           =   7470
   End
   Begin MCI.MMControl mp3 
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   582
      _Version        =   393216
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
   Begin VB.Timer Timer_Zaman 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   735
      Top             =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1155
      Top             =   135
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   5985
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Geçen Süre :"
      Height          =   195
      Left            =   465
      TabIndex        =   13
      Top             =   2880
      Width           =   945
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Zaman :"
      Height          =   195
      Left            =   825
      TabIndex        =   12
      Top             =   2565
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Þarkýnýn zamaný :"
      Height          =   195
      Left            =   210
      TabIndex        =   11
      Top             =   3225
      Width           =   1200
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1575
      TabIndex        =   10
      Top             =   3180
      Width           =   75
   End
   Begin VB.Label Label6 
      Caption         =   $"AntiDsc.frx":1562E
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   210
      TabIndex        =   7
      Top             =   4305
      Width           =   2745
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1575
      TabIndex        =   3
      Top             =   2835
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1575
      TabIndex        =   2
      Top             =   2520
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   4
      Top             =   315
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "AntiDsc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Byte: Dim max_a As Byte: Dim DakSan As String
Private Sub Command1_Click()
    On Error GoTo RenkiPtal
    CDialog.ShowColor
    File1.Enabled = False: mp3.FileName = App.Path + "\mp3\" + File1.List(File1.ListIndex): mp3.Command = "Open"
    Command1.Caption = max_a - a: Timer_Zaman.Enabled = True: Command1.Enabled = False: List1.Enabled = False: Command3.Enabled = False
    List1.Clear: List1.AddItem File1.List(File1.ListIndex) & "," & CStr(CDialog.Color)
RenkiPtal:
    Exit Sub
End Sub
Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 27
        List1.AddItem "0,0," & Label2.Caption
        BirMüzikSonu
    Case 49
        List1.AddItem "1,0," & Label2.Caption
    Case 50
        List1.AddItem "2,0," & Label2.Caption
    Case 51
        List1.AddItem "3,0," & Label2.Caption
    Case 52
        List1.AddItem "3,4," & Label2.Caption
    Case 53
        List1.AddItem "5,0," & Label2.Caption
    Case 97
        List1.AddItem "1,0," & Label2.Caption
    Case 98
        List1.AddItem "2,0," & Label2.Caption
    Case 99
        List1.AddItem "3,0," & Label2.Caption
    Case 100
        List1.AddItem "3,4," & Label2.Caption
    Case 101
        List1.AddItem "5,0," & Label2.Caption
    End Select
End Sub
Private Sub Command3_Click()
    If Dir(App.Path & "\mp3\" & Left(File1.List(File1.ListIndex), Len(File1.List(File1.ListIndex)) - 3) & "dsc") <> "" Then
        If MsgBox("Bu þarkýnýn diskosu daha önceden oluþturulmuþ." + Chr(13) + "Üzerine yazmak istiyormusunuz?", vbExclamation + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
    End If
    Open App.Path & "\mp3\" & Left(File1.List(File1.ListIndex), Len(File1.List(File1.ListIndex)) - 3) & "dsc" For Output As #1
        List1.ListIndex = 0
        Write #1, File1.List(File1.ListIndex), Val(CStr(CDialog.Color))
        Do While List1.ListCount <> List1.ListIndex + 1
            List1.ListIndex = List1.ListIndex + 1
            Write #1, Val(Left(List1.List(List1.ListIndex), 1)), Val(Mid(List1.List(List1.ListIndex), 3, 1)), Val(Mid(List1.List(List1.ListIndex), 5, Len(List1.List(List1.ListIndex)) - 4))
        Loop
    Close #1
End Sub
Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then End
End Sub
Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 27
        BirMüzikSonu
    Case 13
        Command1_Click
    Case 32
        Command1_Click
    Case Else
        If File1.Enabled = True Then List1.Clear: Command3.Enabled = False
    End Select
End Sub
Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If File1.Enabled = True Then List1.Clear: Command3.Enabled = False
End Sub
Private Sub Form_Click()
    BirMüzikSonu
End Sub
Private Sub Form_Load()
    File1.Path = App.Path + "\mp3\": max_a = 3
    File1.Width = Me.Width:  If File1.ListCount = 0 Then File1.Height = 226 Else File1.Height = Val(File1.ListCount) * 226: File1.Selected(0) = True
    If File1.Height > Me.Height / 2 Then File1.Height = Me.Height / 2
End Sub
Sub BirMüzikSonu()
    If File1.Enabled = False Then
        File1.Enabled = True: List1.Enabled = True: Command1.Enabled = True
        Timer_Zaman.Enabled = False: Timer1.Enabled = False: Command2.Enabled = False
        mp3.Command = "Close": Command1.Caption = "Baþla": a = 0
        Label1.Caption = "": Label2.Caption = "": Label8.Caption = ""
        Command3.Enabled = True: Command3.SetFocus
    Else
        End
    End If
End Sub
Private Sub Label1_Change()
    Label2.Caption = Val(Label2.Caption) + 1
End Sub
Private Sub Label2_Change()
    If Len(CStr(Val(Label2.Caption) \ 60)) = 1 Then DakSan = "0" Else DakSan = ""
    Label8.Caption = DakSan & Val(Label2.Caption) \ 60 & ":"
    If Len(CStr(Val(Label2.Caption) Mod 60)) = 1 Then DakSan = "0" Else DakSan = ""
    Label8.Caption = Label8.Caption & DakSan & Val(Label2.Caption) Mod 60
End Sub
Private Sub Label3_Change()
    a = a + 1:: Command1.Caption = max_a - a
    If a = max_a Then
        Command2.Enabled = True: Timer1.Enabled = True
        mp3.Command = "play"
        Timer_Zaman.Enabled = False
        Label1.Caption = Time: Label2.Caption = 0: a = 0
    End If
End Sub
Private Sub mp3_Done(NotifyCode As Integer)
    If NotifyCode = 1 Then List1.AddItem "0,0," & Label2.Caption: BirMüzikSonu
End Sub
Private Sub Timer_Zaman_Timer()
    Label3.Caption = Time
End Sub
Private Sub Timer1_Timer()
    Label1.Caption = Time
End Sub
