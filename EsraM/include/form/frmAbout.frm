VERSION 5.00
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form Hakkýnda 
   BackColor       =   &H0096E06D&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3285
   ClientLeft      =   2310
   ClientTop       =   1515
   ClientWidth     =   6405
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2267.366
   ScaleMode       =   0  'User
   ScaleWidth      =   6014.626
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      Height          =   1740
      Left            =   60
      ScaleHeight     =   1179.92
      ScaleMode       =   0  'User
      ScaleWidth      =   969.22
      TabIndex        =   0
      Top             =   127
      Width           =   1440
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   0
         Picture         =   "frmAbout.frx":000C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1380
      End
   End
   Begin OsenXPCntrl.OsenXPButton cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   495
      Left            =   4785
      TabIndex        =   5
      Top             =   2122
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Tamam"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmAbout.frx":044E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":046A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   1635
      TabIndex        =   4
      Top             =   1117
      Width           =   4575
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ESRAM DÝKÝM EVÝ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   1635
      TabIndex        =   2
      Top             =   127
      Width           =   4245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   56.343
      X2              =   5873.769
      Y1              =   1392.17
      Y2              =   1392.17
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sürüm"
      Height          =   195
      Left            =   2475
      TabIndex        =   3
      Top             =   652
      Width           =   450
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":054D
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   75
      TabIndex        =   1
      Top             =   2122
      Width           =   4830
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Hakkýnda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Unload(Cancel As Integer): Anasayfa.Enabled = True: End Sub
Private Sub Form_KeyPress(KeyAscii As Integer): Unload Me: End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single): Unload Me: End Sub
Private Sub cmdOK_Click(): Unload Me: End Sub
Private Sub Form_Load()
    Dim Control As Control: Me.BackColor = rnk_frm_arka: Me.Move (frmMain.ScaleWidth - Width) / 2, (frmMain.ScaleHeight - Height) / 2
    For Each Control In Me
        If TypeOf Control Is OsenXPButton Then Control.BackColor = rnk_btn_arka: Control.BackOver = rnk_btn_arka: Control.ForeColor = rnk_btn_ön: Control.ForeOver = rnk_btn_ön
        If TypeOf Control Is Label Then Control.ForeColor = rnk_frm_ön
        If TypeOf Control Is Line Then Control.BorderColor = rnk_frm_ön
    Next Control
    Me.Caption = App.Title & " Hakkýnda": lblVersion.Caption = "Sürüm " & App.Major & "." & App.Minor & "." & App.Revision: lblTitle.Caption = App.Title
End Sub
