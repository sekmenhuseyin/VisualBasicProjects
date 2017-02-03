VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form frmOptions 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayarlar"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin OsenXPCntrl.OsenXPButton cmd_ok 
      Default         =   -1  'True
      Height          =   435
      Left            =   3960
      TabIndex        =   2
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
      MICON           =   "frmOptions.frx":15242
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
      TabIndex        =   1
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
      MICON           =   "frmOptions.frx":1525E
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
      TabPicture(0)   =   "frmOptions.frx":1527A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Option5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Mesaj && Uygulama"
      TabPicture(1)   =   "frmOptions.frx":15296
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Bilgisayarý Kapat"
      TabPicture(2)   =   "frmOptions.frx":152B2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame9"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame9 
         Caption         =   "              "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   -74790
         TabIndex        =   14
         Top             =   615
         Width           =   5160
         Begin VB.ComboBox cmb_Tema 
            Height          =   315
            ItemData        =   "frmOptions.frx":152CE
            Left            =   1335
            List            =   "frmOptions.frx":152E1
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1710
            Width           =   3615
         End
         Begin OsenXPCntrl.OsenXPButton cmd_Change1 
            Height          =   330
            Left            =   1755
            TabIndex        =   15
            Top             =   345
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   582
            BTYPE           =   9
            TX              =   "Deðiþtir"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   -2147483633
            BCOLO           =   -2147483633
            FCOL            =   -2147483640
            FCOLO           =   -2147483640
            MCOL            =   12632256
            MPTR            =   0
            MICON           =   "frmOptions.frx":15328
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.OsenXPButton cmd_Change2 
            Height          =   330
            Left            =   1755
            TabIndex        =   17
            Top             =   660
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   582
            BTYPE           =   9
            TX              =   "Deðiþtir"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   -2147483633
            BCOLO           =   -2147483633
            FCOL            =   -2147483640
            FCOLO           =   -2147483640
            MCOL            =   12632256
            MPTR            =   0
            MICON           =   "frmOptions.frx":15344
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.OsenXPButton cmd_Change3 
            Height          =   330
            Left            =   1755
            TabIndex        =   18
            Top             =   975
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   582
            BTYPE           =   9
            TX              =   "Deðiþtir"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   -2147483633
            BCOLO           =   -2147483633
            FCOL            =   -2147483640
            FCOLO           =   -2147483640
            MCOL            =   12632256
            MPTR            =   0
            MICON           =   "frmOptions.frx":15360
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.OsenXPButton cmd_Change4 
            Height          =   330
            Left            =   1755
            TabIndex        =   19
            Top             =   1290
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   582
            BTYPE           =   9
            TX              =   "Deðiþtir"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   -2147483633
            BCOLO           =   -2147483633
            FCOL            =   -2147483640
            FCOLO           =   -2147483640
            MCOL            =   12632256
            MPTR            =   0
            MICON           =   "frmOptions.frx":1537C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lbl_Color1 
            BackColor       =   &H80000003&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1335
            TabIndex        =   29
            Top             =   345
            Width           =   330
         End
         Begin VB.Label lbl_Color2 
            BackColor       =   &H80000008&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1335
            TabIndex        =   28
            Top             =   660
            Width           =   330
         End
         Begin VB.Label lbl_Color3 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1335
            TabIndex        =   27
            Top             =   975
            Width           =   330
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Arkaplan rengi"
            Height          =   195
            Left            =   105
            TabIndex        =   26
            Top             =   420
            Width           =   1020
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Yazý rengi"
            Height          =   195
            Left            =   105
            TabIndex        =   25
            Top             =   735
            Width           =   690
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Düðme rengi"
            Height          =   195
            Left            =   105
            TabIndex        =   24
            Top             =   1050
            Width           =   900
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Temalar"
            Height          =   195
            Left            =   105
            TabIndex        =   23
            Top             =   1785
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Düðme yazýsý"
            Height          =   195
            Left            =   105
            TabIndex        =   22
            Top             =   1358
            Width           =   930
         End
         Begin VB.Label lbl_Color4 
            BackColor       =   &H80000008&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1335
            TabIndex        =   21
            Top             =   1290
            Width           =   330
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Görünüm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   20
            Top             =   0
            Width           =   765
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "                 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   -74790
         TabIndex        =   7
         Top             =   615
         Width           =   5160
         Begin VB.CheckBox Set_Check1 
            Caption         =   "Alarm kurulu olarak baþla"
            Height          =   270
            Left            =   105
            TabIndex        =   12
            Top             =   345
            Width           =   4815
         End
         Begin VB.CheckBox Set_Check2 
            Caption         =   "Simge durumunda baþla"
            Height          =   270
            Left            =   105
            TabIndex        =   11
            Top             =   660
            Width           =   4815
         End
         Begin VB.CheckBox Set_Check3 
            Caption         =   "Alarm kurulduðunda simge durumuna küçült"
            Height          =   270
            Left            =   105
            TabIndex        =   10
            Top             =   975
            Width           =   4815
         End
         Begin VB.CheckBox Set_Check4 
            Caption         =   "Alarm çaldýðýnda önceki durumuna getir"
            Height          =   270
            Left            =   105
            TabIndex        =   9
            Top             =   1290
            Width           =   4815
         End
         Begin VB.CheckBox Set_Check5 
            Caption         =   "Etkinlikten sonra kapan"
            Height          =   270
            Left            =   105
            TabIndex        =   8
            Top             =   1605
            Width           =   4815
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seçenekler"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   13
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.CheckBox Option5 
         Caption         =   "Saat baþlarýnda alarm ver"
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
         Left            =   390
         TabIndex        =   3
         Top             =   550
         Width           =   2580
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   210
         TabIndex        =   4
         Top             =   615
         Width           =   5160
         Begin VB.ComboBox opt_Type 
            Height          =   315
            ItemData        =   "frmOptions.frx":15398
            Left            =   1095
            List            =   "frmOptions.frx":153A5
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   360
            Width           =   3915
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alarm Tipi"
            Height          =   195
            Left            =   255
            TabIndex        =   6
            Top             =   420
            Width           =   690
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
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim combo_DURUM As Boolean
Private Sub cmd_cancel_Click()
    OptionsOn = False 'ayarlar kapalý
    frmAlarm.cmdOptions_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
    cmd_cancel_Click
End Sub
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub cmd_ok_Click()
    'about option5      saat baþlarý
    option_NO5 = Option5.Value
    opt5_opt_Type = opt_Type.ListIndex
    'Settings
    Settings_opt1 = Set_Check1.Value
    Settings_opt2 = Set_Check2.Value
    Settings_opt3 = Set_Check3.Value
    Settings_opt4 = Set_Check4.Value
    Settings_opt5 = Set_Check5.Value
    'Görünüm
    ColorNo1 = CStr(lbl_Color1.BackColor)
    ColorNo2 = CStr(lbl_Color2.BackColor)
    ColorNo3 = CStr(lbl_Color3.BackColor)
    ColorNo4 = CStr(lbl_Color4.BackColor)
    Color_Theme = cmb_Tema.ListIndex
    'exit frmoptions
    Call frmAlarm.AnaGörünüm: Call cmd_cancel_Click
End Sub
Private Sub Form_Load()
    'about option5      saat baþlarý
    Option5.Value = option_NO5
    opt_Type.ListIndex = opt5_opt_Type
    'Settings
    Set_Check1.Value = Settings_opt1
    Set_Check2.Value = Settings_opt2
    Set_Check3.Value = Settings_opt3
    Set_Check4.Value = Settings_opt4
    Set_Check5.Value = Settings_opt5
    'Görünüm
    lbl_Color1.BackColor = ColorNo1: SSTab.BackColor = ColorNo1: Me.BackColor = ColorNo1
    lbl_Color2.BackColor = ColorNo2
    lbl_Color3.BackColor = ColorNo3: cmd_cancel.BackColor = ColorNo3: cmd_cancel.BackOver = ColorNo3: cmd_ok.BackColor = ColorNo3: cmd_ok.BackOver = ColorNo3
    lbl_Color4.BackColor = ColorNo4: cmd_cancel.ForeColor = ColorNo4: cmd_cancel.ForeOver = ColorNo4: cmd_ok.ForeColor = ColorNo4: cmd_ok.ForeOver = ColorNo4
    combo_DURUM = False: cmb_Tema.ListIndex = Color_Theme: combo_DURUM = True
    OptionsOn = True 'ayarlar açýk
End Sub
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub cmb_Tema_Click()
    If combo_DURUM = False Then Exit Sub
    RenkTemalarý (cmb_Tema.List(cmb_Tema.ListIndex))
    lbl_Color1.BackColor = ColorNo1
    lbl_Color2.BackColor = ColorNo2
    lbl_Color3.BackColor = ColorNo3
    lbl_Color4.BackColor = ColorNo4
End Sub
Private Sub cmd_Change1_Click()
    On Error GoTo iptalebasýldý
    CDialog.ShowColor
    lbl_Color1.BackColor = CDialog.Color
    cmb_Tema.ListIndex = 0
    Exit Sub
iptalebasýldý:
End Sub
Private Sub cmd_Change2_Click()
    On Error GoTo iptalebasýldý
    CDialog.ShowColor
    lbl_Color2.BackColor = CDialog.Color
    cmb_Tema.ListIndex = 0
    Exit Sub
iptalebasýldý:
End Sub
Private Sub cmd_Change3_Click()
    On Error GoTo iptalebasýldý
    CDialog.ShowColor
    lbl_Color3.BackColor = CDialog.Color
    cmb_Tema.ListIndex = 0
    Exit Sub
iptalebasýldý:
End Sub
Private Sub cmd_Change4_Click()
    On Error GoTo iptalebasýldý
    CDialog.ShowColor
    lbl_Color4.BackColor = CDialog.Color
    cmb_Tema.ListIndex = 0
    Exit Sub
iptalebasýldý:
End Sub
Private Sub SSTab_Click(PreviousTab As Integer)
    On Error Resume Next
    Select Case SSTab.Tab
        Case 0: opt_Type.SetFocus
        Case 1: cmd_ok.SetFocus
        Case 2: cmd_Change1.SetFocus
    End Select
End Sub
Private Sub Timer1_Timer()
    Me.Left = frmAlarm.Left + frmAlarm.Width
    Me.Top = frmAlarm.Top
End Sub
