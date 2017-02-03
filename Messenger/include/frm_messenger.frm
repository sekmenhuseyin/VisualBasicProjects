VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Object = "{9898A615-558F-4B19-8D75-2D993AD12970}#1.0#0"; "TrayIcon.ocx"
Object = "{1C61EB28-849A-4235-A869-D95B475DD9E7}#1.0#0"; "prjXPToolBar.ocx"
Begin VB.Form frm_Messenger 
   BackColor       =   &H0080C0FF&
   Caption         =   "iso Messenger"
   ClientHeight    =   7500
   ClientLeft      =   4035
   ClientTop       =   2790
   ClientWidth     =   6495
   FillStyle       =   0  'Solid
   Icon            =   "frm_messenger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   6495
   Begin RichTextLib.RichTextBox Text_Giden 
      Height          =   1335
      Left            =   60
      TabIndex        =   0
      Top             =   5130
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   2355
      _Version        =   393217
      BackColor       =   12640511
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frm_messenger.frx":1CFA
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2220
      Top             =   6540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_messenger.frx":1D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_messenger.frx":4C10
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_messenger.frx":9D02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbNetIP 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   6495
   End
   Begin prjXPToolBar.ctlXPToolBar ToolMSN 
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TrayIconOCX.ToolTipOnDemand ToolTipOnDemand1 
      Left            =   1785
      Top             =   6570
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin TrayIconOCX.TrayIcon TrayIcon1 
      Left            =   1365
      Top             =   6570
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.ListBox List_Gelen 
      BackColor       =   &H00C0E0FF&
      Height          =   4140
      IntegralHeight  =   0   'False
      ItemData        =   "frm_messenger.frx":EDF4
      Left            =   60
      List            =   "frm_messenger.frx":EDF6
      TabIndex        =   1
      Top             =   930
      Width           =   6375
   End
   Begin OsenXPCntrl.OsenXPButton Sender 
      Default         =   -1  'True
      Height          =   1335
      Left            =   4980
      TabIndex        =   3
      Top             =   5130
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2355
      BTYPE           =   3
      TX              =   "Gönder"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8438015
      BCOLO           =   8438015
      FCOL            =   -2147483641
      FCOLO           =   -2147483641
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frm_messenger.frx":EDF8
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.StatusBar StatusMSN 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   7170
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "08:01"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   6540
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   420
      Top             =   6570
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   420
      Left            =   840
      TabIndex        =   4
      Top             =   6570
      Visible         =   0   'False
      Width           =   420
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   0   'False
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   741
      _cy             =   741
   End
   Begin VB.Menu mn_Tema 
      Caption         =   "Theme"
      Visible         =   0   'False
      Begin VB.Menu mn_Tema_Sub 
         Caption         =   "Standart"
         Index           =   0
      End
      Begin VB.Menu mn_Tema_Sub 
         Caption         =   "Mavimsi"
         Index           =   1
      End
      Begin VB.Menu mn_Tema_Sub 
         Caption         =   "Kömür Karasý"
         Index           =   2
      End
      Begin VB.Menu mn_Tema_Sub 
         Caption         =   "Windows XP"
         Index           =   3
      End
   End
   Begin VB.Menu mn_PPP 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mn_PPP_Show 
         Caption         =   "Göster"
         Enabled         =   0   'False
      End
      Begin VB.Menu mn_PPP_Hide 
         Caption         =   "Gizle"
      End
      Begin VB.Menu mn_PPP_Tire 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mn_PPP_Exit 
         Caption         =   "Çýkýþ"
      End
   End
End
Attribute VB_Name = "frm_Messenger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GizlendiMi As Boolean: Dim MSNDurum As Boolean
Private Sub Sender_Click()
    On Error GoTo SendError
    If MSNDurum = False Then
        List_Gelen.AddItem "**********  Karþý bilgisayar kapalý !  **********"
    Else
        Open NetPath(cmbNetIP.ListIndex) For Append As #2: Write #2, Time$, MyName, Text_Giden.Text: Close #2
    End If
    If Ayar(4) = "1" Then
        List_Gelen.AddItem "     " & MyName & "  ->  " & cmbNetIP.Text & "  (" & Time$ & ") :"
    Else
        List_Gelen.AddItem "     " & MyName & "  ->  " & cmbNetIP.Text & " :"
    End If
    List_Gelen.AddItem (Text_Giden.Text)
    List_Gelen.Selected(List_Gelen.ListCount - 1) = True
    Text_Giden.Text = "": Text_Giden.SetFocus
    Exit Sub
SendError:
    MsgBox Err.Number & ":" & Err.Description
End Sub
Private Sub Text_Giden_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Text_Giden.Text = Left(Text_Giden.Text, Len(Text_Giden.Text) - 2)
        Sender_Click
    End If
End Sub
Private Sub Timer1_Timer() 'karþý PC açýk mý kapalý mý?
    Dim ECHO As ICMP_ECHO_REPLY
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''start of searching of msn_durum
    Call Ping(cmbNetIP.List(cmbNetIP.ListIndex), ECHO)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''end of searching of msn_durum
    'þimdi ise duruma baðlý olarak gelen msejlara bakýlacak...
    If ECHO.status = 0 Then
        'karþý PCnin açýk olduðunu belirtir
        StatusMSN.Panels.Item(2).Text = "Karþý PC Açýk"
        Timer1.Interval = 1000
        MSNDurum = True
    Else
        'karþý PCnin kapalý olduðunu belirtir
        StatusMSN.Panels.Item(2).Text = "Karþý PC Kapalý"
        Timer1.Interval = 10000
        MSNDurum = False
    End If
End Sub
Private Sub Timer2_Timer() 'bana gelen mesaj var mý?
    Dim MsgTime, MsgFrom, MsgTxt As String: Dim MsgVarMý As Boolean
    If Dir(MyPath) = "" Then GoTo DosyaYok
    'gelen mesaj varmý diye bakýyor.
    Open MyPath For Input As #1: MsgVarMý = False
bas:
        If EOF(1) Then GoTo son 'yoksa yada mesajýn sonuysa...
        'varsa hemen mesajý yazýp sesi mesaj sesini çaldýrtýr
        Input #1, MsgTime, MsgFrom, MsgTxt
        If Ayar(4) = "1" Then
            List_Gelen.AddItem "     " & MsgFrom & "  (" & Time$ & ") :"
        Else
            List_Gelen.AddItem "     " & MsgFrom & " :"
        End If
        List_Gelen.AddItem MsgTxt
        List_Gelen.Selected(List_Gelen.ListCount - 1) = True 'son gelen mesaj seçilir
        MsgTime = "": MsgFrom = "": MsgTxt = "": MsgVarMý = True: GoTo bas
son:
    Close #1
    If MsgVarMý = False Then
        Exit Sub
    Else
        If Ayar(2) = 1 Then WMP.Controls.play 'mesaj uyarý sesi
        Call ShowBalloon(TrayIcon1, blIconInfo, MsgFrom & " size mesaj yazdý!", Me.Caption)
    End If
DosyaYok:
    Open MyPath For Output As #1: Close #1
End Sub
Private Sub cmbNetIP_Click()
    Call Timer1_Timer
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub mn_PPP_Exit_Click()
    Unload Me
End Sub
Private Sub mn_PPP_Hide_Click()
    Me.Hide: GizlendiMi = True
    mn_PPP_Hide.Enabled = False: mn_PPP_Show.Enabled = True
    Call ShowBalloon(TrayIcon1, blIconInfo, "iso MSN hala çalýþýyor !", Me.Caption)
End Sub
Private Sub mn_PPP_Show_Click()
    Me.Show: If Konum(2) = "0" Then Me.WindowState = 0 Else Me.WindowState = 2: Me.Text_Giden.SetFocus: GizlendiMi = False
    If Me.Left > Screen.Width Then Me.Left = Screen.Width - Me.Width
    If Me.Top > Screen.Height Then Me.Top = Screen.Height - Me.Height
    mn_PPP_Hide.Enabled = True: mn_PPP_Show.Enabled = False
End Sub
Private Sub ToolMSN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mn_Tema.Visible = False: ToolMSN.Highlight "btnTheme", 0
End Sub
Private Sub ToolMSN_ChildMouseDown(sKey As String, Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case sKey
        Case "btnTheme": ToolMSN.Highlight "btnTheme", 0
        Case "btnOptions": Me.Enabled = False: frm_Options.Show , Me
        Case "btnExit": Unload Me
    End Select
End Sub
Private Sub ToolMSN_ChildMouseUp(sKey As String, Button As Integer, Shift As Integer, x As Single, y As Single)
    If sKey = "btnTheme" And mn_Tema.Visible = False Then PopupMenu mn_Tema, , ToolMSN.Buttons("btnTheme").Left, _
        ToolMSN.Buttons("btnTheme").Top + ToolMSN.Buttons("btnTheme").Height
    ToolMSN.Highlight "btnTheme", 0
End Sub
Private Sub mn_Tema_Sub_Click(index As Integer)
    Dim i As Byte: For i = 0 To 4: Boya(i) = RenkTemalarý(mn_Tema_Sub(index).Caption, i): Next i
    BoyaTheme = mn_Tema_Sub(index).Caption: Call AnaGörünüm: ToolMSN.Highlight "btnTheme", 0
End Sub
'startof form codes *+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Public Sub AnaGörünüm()
    Write #7, Time$, Me.Name, "AnaGörünüm", "Start" 'logging
    On Error Resume Next: Me.BackColor = Boya(0): WMP.URL = Ayar(3)
    Text_Giden.BackColor = Boya(1): Text_Giden.SelStart = 0: Text_Giden.SelLength = Len(Text_Giden.Text): Text_Giden.SelColor = Boya(2): Text_Giden.SelLength = 0: Text_Giden.SetFocus
    List_Gelen.BackColor = Boya(1): List_Gelen.ForeColor = Boya(2)
    Sender.BackColor = Boya(3): Sender.BackOver = Boya(3): Sender.ForeColor = Boya(4): Sender.ForeOver = Boya(4)
    cmbNetIP.Clear: Dim i As Byte
    For i = 0 To 254
        If NetName(i) <> "" Then cmbNetIP.AddItem NetName(i) Else Exit For
    Next i
    If cmbNetIP.ListCount <> 0 Then cmbNetIP.ListIndex = 0
    For i = 0 To 3
        If mn_Tema_Sub(i).Caption = BoyaTheme Then mn_Tema_Sub(i).Checked = True Else mn_Tema_Sub(i).Checked = False
    Next i
    Write #7, Time$, Me.Name, "AnaGörünüm", "End" 'logging
End Sub
Private Sub Form_Load()
    Write #7, Time$, Me.Name, "Form_Load", "Start" 'logging
    If Konum(2) = 0 Then Me.Move Konum(0), Konum(1), Kenar(0), Kenar(1) Else Me.WindowState = 2: Me.Width = Kenar(0): Me.Height = Kenar(1)
    With TrayIcon1: .IconHandle = Me.Icon: .toolTip = Me.Caption: .Create Me.hwnd: End With
    With ToolMSN
        .Add "btnTheme"
            .Buttons("btnTheme").Style = 3
            .Buttons("btnTheme").Picture = ImageList1.ListImages(1).Picture
            .Buttons("btnTheme").Caption = "Tema"
            .Highlight "btnTheme", 0
        .Add "btnOptions"
            .Buttons("btnOptions").Style = 3
            .Buttons("btnOptions").Picture = ImageList1.ListImages(2).Picture
            .Buttons("btnOptions").Caption = "Seçenekler"
            .Highlight "btnOptions", 0
        .Add "btnExit"
            .Buttons("btnExit").Style = 0
            .Buttons("btnExit").Picture = ImageList1.ListImages(3).Picture
            .Buttons("btnExit").Caption = "Çýkýþ"
            .Highlight "btnExit", 0
    End With
    DockingStart Me, True: GizlendiMi = False: Call AnaGörünüm
    If Ayar(1) = "1" Then Me.WindowState = 1
    If Ayar(5) = "1" Then AlwaysOnTop Me, True
    Write #7, Time$, Me.Name, "Form_Load", "Successful" 'logging
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27: WindowState = 1
    End Select
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mn_PPP.Visible = False
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mn_PPP
End Sub
Private Sub Form_Resize()
    If Me.WindowState = 1 And GizlendiMi = False Then
        Call mn_PPP_Hide_Click: Exit Sub
    ElseIf Me.WindowState = 1 Then: Exit Sub
    ElseIf Me.WindowState = 2 Then: Konum(2) = 2
    ElseIf Me.WindowState = 0 Then
        If Width < 5000 Then Width = 5000
        If Height < 8000 Then Height = 8000
        Kenar(0) = Me.Width: Kenar(1) = Me.Height: Konum(2) = 0
    End If
    'component movements
    ToolMSN.Move 0, 0, ScaleWidth
    cmbNetIP.Move 0, ToolMSN.Height - 40, ScaleWidth
    List_Gelen.Move 40, cmbNetIP.Top + cmbNetIP.Height + 40, ScaleWidth - 80, ScaleHeight - Text_Giden.Height - ToolMSN.Height - cmbNetIP.Height - StatusMSN.Height - 200
    Text_Giden.Move 40, List_Gelen.Top + List_Gelen.Height + 100, ScaleWidth - Sender.Width - 120
    Sender.Move Text_Giden.Width + 80, Text_Giden.Top
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Write #7, Time$, Me.Name, "Form_Unload", "Start" 'logging
    If Me.WindowState = 0 Then
        If Me.Left > Screen.Width Then Konum(0) = Screen.Width - Me.Width Else Konum(0) = Me.Left
        If Me.Top > Screen.Height Then Konum(1) = Screen.Height - Me.Height Else Konum(1) = Me.Top
        Kenar(0) = Me.Width: Kenar(1) = Me.Height
    End If
    TrayIcon1.Remove: ToolTipOnDemand1.Destroy ': DockingTerminate Me
    Write #7, Time$, Me.Name, "Form_Unload", "Successful" 'logging
    TheEnd
End Sub
'endof form codes *+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
'startof trayicon codes *+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
Private Sub ToolTipOnDemand1_BalloonDestroyed()
    TrayIcon1.TrackIconMovement = False
End Sub
Private Sub ToolTipOnDemand1_BalloonShowed()
    TrayIcon1.TrackIconMovement = True
End Sub
Private Sub TrayIcon1_TrayMouseEvent(ByVal MouseEvent As stMouseEvent)
    Select Case MouseEvent
        Case stMouseMove
        Case stLeftButtonDown
            mn_PPP.Visible = False
        Case stLeftButtonUp
            Call mn_PPP_Show_Click
        Case stLeftButtonDoubleClick
        Case stRightButtonDown
            mn_PPP.Visible = False
        Case stRightButtonUp
            PopupMenu mn_PPP
        Case stRightButtonDoubleClick
        Case stMiddleButtonDown
        Case stMiddleButtonUp
        Case stMiddleButtonDoubleClick
    End Select
End Sub
Private Sub ToolTipOnDemand1_MouseEvents(MouseEvent As Long)
    Select Case MouseEvent
        Case stMouseMove
        Case stLeftButtonDown
            ToolTipOnDemand1.Destroy
        Case stLeftButtonUp
        Case stLeftButtonDoubleClick
        Case stRightButtonDown
            ToolTipOnDemand1.Destroy
        Case stRightButtonUp
        Case stRightButtonDoubleClick
        Case stMiddleButtonDown
        Case stMiddleButtonUp
        Case stMiddleButtonDoubleClick
    End Select
End Sub
Private Sub ShowBalloon(ByVal SystemTrayIcon As TrayIcon, ByVal enIconType As blIconType, ByVal sPrompt As String, Optional ByVal sTitle As String, _
                        Optional ByVal lTimeout As Long = 3000, Optional ByVal lBackColor As Long = -1, Optional ByVal lForeColor As Long = -1)
    Dim lX As Long, lY As Long
    Call SystemTrayIcon.GetIconMiddle(lX, lY)
    If lForeColor = -1 Then lForeColor = vbBlack
    If lBackColor = -1 Then lBackColor = &H80000018
    With ToolTipOnDemand1
        .ParentHwnd = SystemTrayIcon.SysTrayHWnd
        .x = lX
        .y = lY
        .BackColor = lBackColor
        .ForeColor = lForeColor
        .Prompt = sPrompt
        .Title = sTitle
        .Timeout = lTimeout
        .IconType = enIconType
        .Show
      End With
End Sub
'endof trayicon codes *+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*


