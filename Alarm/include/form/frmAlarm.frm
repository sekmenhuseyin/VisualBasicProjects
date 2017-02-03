VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAlarm 
   BackColor       =   &H007E6624&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alarm"
   ClientHeight    =   6150
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4320
   Icon            =   "frmAlarm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   3600
      Top             =   4080
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   240
      TabIndex        =   19
      Top             =   2880
      Width           =   3375
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   17
      Top             =   5715
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin OsenXPCntrl.OsenXPButton cmdOptions 
      Height          =   390
      Left            =   1890
      TabIndex        =   2
      Top             =   105
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   688
      BTYPE           =   3
      TX              =   ">>  >>  >>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483645
      BCOLO           =   -2147483645
      FCOL            =   8283684
      FCOLO           =   8283684
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmAlarm.frx":15162
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   480
      Top             =   6045
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   840
      Top             =   6045
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H007E6624&
      Caption         =   "         "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Left            =   105
      TabIndex        =   5
      Top             =   525
      Width           =   4110
      Begin VB.ListBox List1 
         Height          =   2895
         IntegralHeight  =   0   'False
         ItemData        =   "frmAlarm.frx":1517E
         Left            =   105
         List            =   "frmAlarm.frx":15180
         TabIndex        =   11
         Top             =   1365
         Width           =   3375
      End
      Begin VB.ComboBox dakka 
         BackColor       =   &H007E6624&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmAlarm.frx":15182
         Left            =   3360
         List            =   "frmAlarm.frx":1523A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   585
      End
      Begin VB.ComboBox saat 
         BackColor       =   &H007E6624&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmAlarm.frx":1532E
         Left            =   2625
         List            =   "frmAlarm.frx":1537A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   585
      End
      Begin OsenXPCntrl.OsenXPButton cmdAdd 
         Height          =   390
         Left            =   105
         TabIndex        =   12
         Top             =   840
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   688
         BTYPE           =   3
         TX              =   "Ekle"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   8283684
         FCOLO           =   8283684
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmAlarm.frx":153DE
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdDel 
         Height          =   390
         Left            =   2835
         TabIndex        =   13
         Top             =   840
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   688
         BTYPE           =   3
         TX              =   "Sil"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   8283684
         FCOLO           =   8283684
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmAlarm.frx":153FA
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdSettings 
         Height          =   2910
         Left            =   3570
         TabIndex        =   14
         Top             =   1365
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   5133
         BTYPE           =   3
         TX              =   ">>  >>  >>"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   8283684
         FCOLO           =   8283684
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmAlarm.frx":15416
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker tarih 
         Height          =   330
         Left            =   105
         TabIndex        =   15
         Top             =   420
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   20971521
         CurrentDate     =   39155
      End
      Begin OsenXPCntrl.OsenXPButton cmdEdit 
         Height          =   390
         Left            =   1360
         TabIndex        =   18
         Top             =   840
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   688
         BTYPE           =   3
         TX              =   "Deðiþtir"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483645
         BCOLO           =   -2147483645
         FCOL            =   8283684
         FCOLO           =   8283684
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmAlarm.frx":15432
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lbl_tarih 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tarih"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   16
         Top             =   210
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "Alarm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lbl_dakika 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dakika"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3360
         TabIndex        =   8
         Top             =   210
         Width           =   510
      End
      Begin VB.Label lbl_saat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saat"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2625
         TabIndex        =   7
         Top             =   210
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   3225
         TabIndex        =   6
         Top             =   315
         Width           =   120
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   105
      Top             =   6045
   End
   Begin OsenXPCntrl.OsenXPButton Command1 
      Default         =   -1  'True
      Height          =   540
      Left            =   105
      TabIndex        =   10
      Top             =   5040
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   953
      BTYPE           =   3
      TX              =   "Alarmý Kur"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483645
      BCOLO           =   -2147483645
      FCOL            =   8283684
      FCOLO           =   8283684
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmAlarm.frx":1544E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   420
      Left            =   1320
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6045
      Visible         =   0   'False
      Width           =   420
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   5
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   -1  'True
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   1725
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mn1 
         Caption         =   "&Göster"
         Enabled         =   0   'False
      End
      Begin VB.Menu mn2 
         Caption         =   "&Sakla"
      End
      Begin VB.Menu mn3 
         Caption         =   "-"
      End
      Begin VB.Menu mn8 
         Caption         =   "Alarmý Kur"
      End
      Begin VB.Menu mn7 
         Caption         =   "Alarmý Kapat"
         Enabled         =   0   'False
      End
      Begin VB.Menu mn6 
         Caption         =   "Alarmý Sustur"
         Enabled         =   0   'False
      End
      Begin VB.Menu mn5 
         Caption         =   "-"
      End
      Begin VB.Menu mn4 
         Caption         =   "&Kapat"
      End
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAdd_Click()
    AlarmCount = List1.ListCount: ReDim Preserve AlarmSettings(AlarmCount)
    AlarmSettings(AlarmCount).AlarmDate = tarih.Value
    AlarmSettings(AlarmCount).AlarmSaat = saat.Text
    AlarmSettings(AlarmCount).AlarmDakka = dakka.Text
    AlarmSettings(AlarmCount).AppOn = False
    AlarmSettings(AlarmCount).MsgOn = False
    AlarmSettings(AlarmCount).ShutOn = False
    AlarmSettings(AlarmCount).SndOn = False
    List1.AddItem tarih.Value & " - " & saat.Text & ":" & dakka.Text
End Sub
Private Sub Command1_Click()
    If Command1.Caption = "Alarmý Kur" Then
            Timer1.Enabled = True
            If Settings_opt3 = 1 Then mn2_Click
            Command1.Caption = "Alarmý Kapat": mn8.Enabled = False: mn7.Enabled = True: mn6.Enabled = False
    ElseIf Command1.Caption = "Alarmý Kapat" Then
            Timer1.Enabled = False
            Command1.Caption = "Alarmý Kur": mn8.Enabled = True: mn7.Enabled = False: mn6.Enabled = False
    ElseIf Command1.Caption = "Alarmý Sustur" Then
            WMP.Controls.stop
            Timer1.Enabled = False
            Command1.Caption = "Alarmý Kur": mn8.Enabled = True: mn7.Enabled = False: mn6.Enabled = False
    End If
    Show_SYStray ("deðiþtir")
    Command1.ToolTipText = Command1.Caption
End Sub
Public Sub cmdSettings_Click()
    If frmOptions.Visible = False Then
        cmdSettings.Caption = "<<  <<  <<"
        frmSettings.Left = Me.Left + Me.Width: frmSettings.Top = Me.Top: frmSettings.Show
    Else
        cmdSettings.Caption = ">>  >>  >>"
        Unload frmOptions
    End If
End Sub
Public Sub cmdOptions_Click()
    If frmOptions.Visible = False Then
        cmdOptions.Caption = "<<  <<  <<"
        frmOptions.Left = Me.Left + Me.Width: frmOptions.Top = Me.Top: frmOptions.Show , Me
    Else
        cmdOptions.Caption = ">>  >>  >>"
        Unload frmOptions
    End If
End Sub
Private Sub Timer1_Timer()
    If saat.ListIndex = Val(Left(Time, 2)) And dakka.ListIndex = Val(Mid(Time, 4, 2)) Then
        If option_NO1 = 1 Then          'ses çal
                alarm_SesÇal
                Command1.Caption = "Alarmý Sustur": mn8.Enabled = False: mn7.Enabled = False: mn6.Enabled = True
        End If
        If option_NO2 = 1 Then      'uygulama çlýþtýr
                alarm_UygulamaÇalýþtýr
                If option_NO1 <> 1 Then Command1.Caption = "Alarmý Kur": mn8.Enabled = True: mn7.Enabled = False: mn6.Enabled = False
        End If
        If option_NO3 = 1 Then      'mesaj ver
                alarm_MesajVer
                If option_NO1 <> 1 Then Command1.Caption = "Alarmý Kur": mn8.Enabled = True: mn7.Enabled = False: mn6.Enabled = False
        End If
        If option_NO4 = 1 Then      'bilgisayarý kapa
                alarm_BilgisayarýKapat
                If option_NO1 <> 1 Then Command1.Caption = "Alarmý Kur": mn8.Enabled = True: mn7.Enabled = False: mn6.Enabled = False
        End If
        If Settings_opt4 = 1 Then mn1_Click
        If Settings_opt5 = 1 And option_NO1 <> 1 Then Unload Me
        Timer1.Enabled = False
    End If
End Sub
Private Sub Timer2_Timer()
    Label1.Caption = Time
    If Timer1.Enabled = True Then
        If Label2.Caption = "" Then Label2.Caption = ":" Else Label2.Caption = ""
    Else
        Label2.Caption = ":"
    End If
    If option_NO5 = 1 Then alarm_SaatBaþý 'saat baþlarýnda alarm ver
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Show_SYStray ("sil")
    'DockingTerminate Me
    The_End
End Sub
Public Sub AnaGörünüm()
    Dim Control As Control
    Me.BackColor = ColorNo1: Frame1.BackColor = ColorNo1
    For Each Control In Me
        If TypeOf Control Is OsenXPButton Then Control.BackColor = ColorNo3: Control.BackOver = ColorNo3: Control.ForeColor = ColorNo4: Control.ForeOver = ColorNo4
        If TypeOf Control Is Label Then Control.ForeColor = ColorNo2
        If TypeOf Control Is ComboBox Then Control.BackColor = ColorNo1: Control.ForeColor = ColorNo2
    Next Control
End Sub
Private Sub Form_Load()
    Dim i As Byte: Dim temp As String
    Call AnaGörünüm: Me.Move temp_X, temp_Y: OptionsOn = False: SettingsOn = False
    Label1.Caption = Time: saat.ListIndex = Val(Time_saat): dakka.ListIndex = Val(Time_dakka)
    gizlendiMi = False: Show_SYStray ("ekle") ': DockingStart Me, True
    Me.Show: If Settings_opt1 = 1 Then Command1_Click
    
    ReDim AlarmSettings(0)
End Sub
Private Sub Form_Resize()
    If OptionsOn = True Then cmdOptions_Click
    If SettingsOn = True Then cmdSettings_Click
    If Me.WindowState = 1 And gizlendiMi = False Then Me.WindowState = 0: Me.Left = temp_X: Me.Top = temp_Y: mn2_Click
End Sub
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown Button, Shift, x, y
End Sub
Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseUp Button, Shift, x, y
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    menu.Visible = False
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Me.PopupMenu menu
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static MSG As Long
    MSG = x / Screen.TwipsPerPixelX
    ' Captura cada evento de botones del Raton
    Select Case MSG
        Case WM_LBUTTONDBLCLK  ' Doble click Boton Izquierdo
        Case WM_LBUTTONDOWN  ' Boton Izquierdo pulsado
        Case WM_LBUTTONUP   ' Boton Izquierdo Soltado
            mn1_Click
        Case WM_RBUTTONDBLCLK ' Doble Click Boton Derecho
        Case WM_RBUTTONDOWN ' Boton derecho pulsado
        Case WM_RBUTTONUP  ' Boton derecho Arriba
            Me.PopupMenu Me.menu
    End Select
    DoEvents
End Sub
Private Sub mn1_Click() 'Göster
    Me.Show: Me.WindowState = 0: gizlendiMi = False
    mn1.Enabled = False: mn2.Enabled = True
End Sub
Private Sub mn2_Click() 'Sakla
    Me.Hide: gizlendiMi = True
    mn1.Enabled = True:  mn2.Enabled = False
End Sub
Private Sub mn4_Click() 'Çýk
    Unload Me
End Sub
Private Sub mn8_Click() 'Alarmý Kur
    Command1_Click
End Sub
Private Sub mn7_Click() 'Alarmý Kapat
    Command1_Click
End Sub
Private Sub mn6_Click() 'Alarmý Sustur
    Command1_Click
End Sub
Private Sub Show_SYStray(SYStrayTipi As String)
    Dim sSysTrayText As String
    Select Case Command1.Caption
        Case "Alarmý Kur"
            sSysTrayText = "Alarm Kapalý !"
        Case "Alarmý Kapat"
            sSysTrayText = "Alarm Açýk !" + Chr(13) + "Kurulduðu saat: " + saat.Text + ":" + dakka.Text + ""
        Case "Alarmý Sustur"
            sSysTrayText = "Alarm Çalýyor !"
    End Select
    ColocarIcono Me.hwnd, Me.Icon.Handle, sSysTrayText, SYStrayTipi
End Sub
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Timer3_Timer()
    If Settings_opt2 = 1 Then mn2_Click
    Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
    List2.Clear: On Error Resume Next
    Dim i As Byte: For i = 0 To AlarmCount
        List2.AddItem AlarmSettings(i).AlarmDate & AlarmSettings(i).AlarmSaat & AlarmSettings(i).AlarmDakka & AlarmSettings(i).AppOn & AlarmSettings(i).MsgOn & AlarmSettings(i).ShutOn & AlarmSettings(i).SndOn
    Next i
End Sub
