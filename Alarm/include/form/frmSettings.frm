VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form frmSettings 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayarlar"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin OsenXPCntrl.OsenXPButton cmd_ok 
      Default         =   -1  'True
      Height          =   435
      Left            =   3960
      TabIndex        =   22
      Top             =   3285
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Tamam"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman TUR"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   -2147483640
      FCOLO           =   -2147483640
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmSettings.frx":15242
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmd_cancel 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   75
      TabIndex        =   21
      Top             =   3285
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Ýptal"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman TUR"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   -2147483633
      BCOLO           =   -2147483633
      FCOL            =   -2147483640
      FCOLO           =   -2147483640
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmSettings.frx":1525E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   3030
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   5345
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   -2147483645
      TabCaption(0)   =   "Ses Çal"
      TabPicture(0)   =   "frmSettings.frx":1527A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Option1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Mesaj && Uygulama"
      TabPicture(1)   =   "frmSettings.frx":15296
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Option3"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "Option2"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Bilgisayarý Kapat"
      TabPicture(2)   =   "frmSettings.frx":152B2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "Option4"
      Tab(2).Control(2)=   "Frame8"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame8 
         Caption         =   "Kapatma Ayarlarý"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74790
         TabIndex        =   14
         Top             =   1500
         Width           =   5160
         Begin VB.CheckBox opt_Force 
            Caption         =   "Zorla"
            Height          =   255
            Left            =   105
            TabIndex        =   15
            Top             =   398
            Width           =   735
         End
         Begin ComCtl2.UpDown txt_Time 
            Height          =   330
            Left            =   3675
            TabIndex        =   16
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   327681
            BuddyControl    =   "Text2"
            BuddyDispid     =   196661
            OrigLeft        =   3720
            OrigTop         =   360
            OrigRight       =   3975
            OrigBottom      =   690
            Max             =   100
            Min             =   -1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   1745027075
            Enabled         =   -1  'True
         End
         Begin OsenXPCntrl.OsenXPText txt_Shutdown_Msg 
            Height          =   375
            Left            =   1335
            TabIndex        =   27
            Top             =   840
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   661
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
         End
         Begin OsenXPCntrl.OsenXPText Text2 
            Height          =   330
            Left            =   3240
            TabIndex        =   28
            Top             =   360
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   582
            Alignment       =   2
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "30"
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Kapanmadan önce"
            Height          =   255
            Left            =   1800
            TabIndex        =   20
            Top             =   398
            Width           =   1455
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kapanma mesajý"
            Height          =   195
            Left            =   105
            TabIndex        =   18
            Top             =   930
            Width           =   1155
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "saniye bekle"
            Height          =   195
            Left            =   4065
            TabIndex        =   17
            Top             =   435
            Width           =   885
         End
      End
      Begin VB.CheckBox Option2 
         Caption         =   "Uygulama Çalýþtýr"
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
         Left            =   -74640
         TabIndex        =   13
         Top             =   540
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Height          =   1260
         Left            =   -74790
         TabIndex        =   12
         Top             =   615
         Width           =   5160
         Begin VB.TextBox txt_Program 
            Height          =   750
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   315
            Width           =   3900
         End
         Begin OsenXPCntrl.OsenXPButton search2 
            Height          =   750
            Left            =   105
            TabIndex        =   24
            Top             =   315
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1323
            BTYPE           =   3
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   -2147483633
            BCOLO           =   -2147483633
            FCOL            =   -2147483640
            FCOLO           =   -2147483640
            MCOL            =   12632256
            MPTR            =   0
            MICON           =   "frmSettings.frx":152CE
            PICN            =   "frmSettings.frx":152EA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   4
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.CheckBox Option4 
         Caption         =   "Bilgisayarý Kapat"
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
         Left            =   -74640
         TabIndex        =   11
         Top             =   540
         Width           =   1815
      End
      Begin VB.CheckBox Option3 
         Caption         =   "Mesaj Ver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   10
         Top             =   1980
         Width           =   1215
      End
      Begin VB.CheckBox Option1 
         Caption         =   "Ses Çal"
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
         Left            =   360
         TabIndex        =   9
         Top             =   540
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Caption         =   "Ses Ayarlarý"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   210
         TabIndex        =   5
         Top             =   1980
         Width           =   5160
         Begin ComCtl2.UpDown txt_Repeat 
            Height          =   330
            Left            =   1545
            TabIndex        =   19
            Top             =   345
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   327681
            BuddyControl    =   "Text1"
            BuddyDispid     =   196665
            OrigLeft        =   1590
            OrigTop         =   345
            OrigRight       =   1845
            OrigBottom      =   675
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   1745027075
            Enabled         =   -1  'True
         End
         Begin MSComctlLib.Slider sld_Volume 
            Height          =   255
            Left            =   2685
            TabIndex        =   6
            Top             =   390
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   450
            _Version        =   393216
            Max             =   100
            SelStart        =   100
            TickFrequency   =   5
            Value           =   100
         End
         Begin OsenXPCntrl.OsenXPText Text1 
            Height          =   330
            Left            =   1215
            TabIndex        =   25
            Top             =   345
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   582
            Alignment       =   2
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "1"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ses"
            Height          =   195
            Left            =   2265
            TabIndex        =   8
            Top             =   420
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tekrar sayýsý"
            Height          =   195
            Left            =   105
            TabIndex        =   7
            Top             =   420
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1260
         Left            =   210
         TabIndex        =   1
         Top             =   615
         Width           =   5160
         Begin OsenXPCntrl.OsenXPButton search1 
            Height          =   750
            Left            =   105
            TabIndex        =   23
            Top             =   315
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1323
            BTYPE           =   3
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   -2147483633
            BCOLO           =   -2147483633
            FCOL            =   -2147483640
            FCOLO           =   -2147483640
            MCOL            =   12632256
            MPTR            =   0
            MICON           =   "frmSettings.frx":2BCAC
            PICN            =   "frmSettings.frx":2BCC8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   4
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txt_Music 
            Height          =   750
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   315
            Width           =   3900
         End
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   -74790
         TabIndex        =   4
         Top             =   1980
         Width           =   5160
         Begin OsenXPCntrl.OsenXPText txt_Message 
            Height          =   330
            Left            =   105
            TabIndex        =   26
            Top             =   315
            Width           =   4860
            _ExtentX        =   8573
            _ExtentY        =   582
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   -74790
         TabIndex        =   2
         Top             =   615
         Width           =   5160
         Begin VB.ComboBox cmb_Shutdown 
            Height          =   315
            ItemData        =   "frmSettings.frx":4268A
            Left            =   105
            List            =   "frmSettings.frx":42697
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   315
            Width           =   4860
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   607
      Top             =   5745
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   75
      Top             =   5745
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   2
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim combo_DURUM As Boolean
Private Sub cmd_cancel_Click()
    SettingsOn = False 'ayarlar kapalý
    frmAlarm.cmdSettings_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
    cmd_cancel_Click
End Sub
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub cmd_ok_Click()
    If Dir(txt_Music.Text) = "" Then
        If MsgBox("Yazdýðýnýz ses dosyasý yolu geçerli deðil !" + Chr(13) + "Yine de devam etmek istiyor musunuz?", vbYesNo, "Dosya Bulunamadý") = vbNo Then Exit Sub
    End If
    'about option1      ses çal
    option_NO1 = Option1.Value
    opt1_txt_Music = txt_Music.Text
    opt1_txt_Repeat = txt_Repeat.Value
    opt1_sld_Volume = sld_Volume.Value
    'about option2      uygulama çalýþrýt
    option_NO2 = Option2.Value
    opt2_txt_Program = txt_Program.Text
    'about option3      mesaj
    option_NO3 = Option3.Value
    opt3_txt_Message = txt_Message.Text
    'about option4      bilgisayarý kapat
    option_NO4 = Option4.Value
    opt4_cmb_Shutdown = cmb_Shutdown.ListIndex
    opt4_opt_Force = opt_Force.Value
    opt4_txt_Time = txt_Time
    opt4_txt_Shutdown_Msg = txt_Shutdown_Msg.Text
    'exit frmoptions
    Call frmAlarm.AnaGörünüm: Call cmd_cancel_Click
End Sub
Private Sub Form_Load()
    'about option1      ses çal
    Option1.Value = option_NO1
    txt_Music.Text = opt1_txt_Music
    txt_Repeat.Value = opt1_txt_Repeat
    sld_Volume.Value = opt1_sld_Volume
    'about option2      uygulama çalýþrýt
    Option2.Value = option_NO2
    txt_Program.Text = opt2_txt_Program
    'about option3      mesaj
    Option3.Value = option_NO3
    txt_Message.Text = opt3_txt_Message
    'about option4      bilgisayarý kapat
    Option4.Value = option_NO4
    cmb_Shutdown.ListIndex = opt4_cmb_Shutdown
    opt_Force.Value = opt4_opt_Force
    txt_Time = opt4_txt_Time
    txt_Shutdown_Msg.Text = opt4_txt_Shutdown_Msg
    combo_DURUM = False: combo_DURUM = True
    SettingsOn = True 'ayarlar açýk
End Sub
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Option1_Click()
    If Option1.Value = 1 And Trim(txt_Music.Text) <> "" Then Call search1_Click
End Sub
Private Sub Option2_Click()
    If Option2.Value = 1 And Trim(txt_Program.Text) <> "" Then Call search2_Click
End Sub
Private Sub search1_Click()
    On Error GoTo iptalebasýldý
    CDialog.Filter = "Ses Dosyalarý (*.wav;*.mp3;*.wma)|*.wav;*.mp3;*.wma|TümDosyalar|*.*"
    CDialog.FileName = App.Path + "\include\sound\horoz.wav"
    CDialog.ShowOpen
    If CDialog.FileName <> "" Then txt_Music.Text = CDialog.FileName
    Exit Sub
iptalebasýldý:
End Sub
Private Sub search2_Click()
    On Error GoTo iptalebasýldý
    On Error Resume Next
    CDialog.Filter = "Çalýþtýrýlabilir Dosyalar|*.exe;*.bat;*.com;*.vbs"
    CDialog.FileName = txt_Program.Text
    CDialog.ShowOpen
    If CDialog.FileName <> "" Then txt_Program.Text = CDialog.FileName
    Exit Sub
iptalebasýldý:
End Sub
Private Sub SSTab_Click(PreviousTab As Integer)
    On Error Resume Next
    Select Case SSTab.Tab
        Case 0: search1.SetFocus
        Case 1: search2.SetFocus
        Case 2: cmb_Shutdown.SetFocus
    End Select
End Sub
Private Sub Timer1_Timer()
    Me.Left = frmAlarm.Left + frmAlarm.Width
    Me.Top = frmAlarm.Top
End Sub
Private Sub txt_Repeat_Change()
    If txt_Repeat.Value = 100 Then
        txt_Repeat = 1
    ElseIf txt_Repeat.Value = 0 Then
        txt_Repeat.Value = 99
    End If
End Sub
Private Sub txt_Time_Change()
    If txt_Time.Value = 100 Then
        txt_Time = 0
    ElseIf txt_Time.Value = -1 Then
        txt_Time.Value = 99
    End If
End Sub
