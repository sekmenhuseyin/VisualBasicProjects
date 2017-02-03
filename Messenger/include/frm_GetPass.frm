VERSION 5.00
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form frm_GetPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gir"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5370
   Icon            =   "frm_GetPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1380
      Width           =   5115
   End
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3780
      TabIndex        =   2
      Top             =   600
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Ýptal"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   -2147483643
      BCOLO           =   -2147483643
      FCOL            =   -2147483640
      FCOLO           =   -2147483640
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frm_GetPass.frx":15162
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdOK 
      Default         =   -1  'True
      Height          =   315
      Left            =   3780
      TabIndex        =   1
      Top             =   180
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Tamam"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   -2147483643
      BCOLO           =   -2147483643
      FCOL            =   -2147483640
      FCOLO           =   -2147483640
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frm_GetPass.frx":1517E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Eski Kodunuz"
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
      Left            =   270
      TabIndex        =   3
      Top             =   270
      Width           =   1170
   End
End
Attribute VB_Name = "frm_GetPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DenemeNo As Byte
Private Sub cmdOK_Click()
    Write #7, Time$, Me.Name, "cmdOK_Click", "Start" 'logging
    Dim DeðiþecekKod As String
    If Güven(1) = "1" Then DeðiþecekKod = Güven(5) Else DeðiþecekKod = Güven(4)
    If CStr(CalculateMD5(txtPass.Text)) = CStr(DeðiþecekKod) Then
        Write #7, Time$, Me.Name, "cmdOK_Click", "PassRight" 'logging
        Güven(1) = "0": frm_Messenger.Show: Unload Me
    Else
        Write #7, Time$, Me.Name, "cmdOK_Click", "PassWrong" 'logging
        MsgBox Caption & "nuz hatalý !", , App.FileDescription
        DenemeNo = DenemeNo + 1
        If Güven(2) = "1" And Val(Güven(3)) <= Val(DenemeNo) And Güven(1) = "0" Then
            Write #7, Time$, Me.Name, "cmdOK_Click", "Locked" 'logging
            MsgBox CStr(DenemeNo) & " defa açma kodunu yanlýþ girdiniz !" & vbCrLf & "Lütfen güvenlik kodunu giriniz !", , App.FileDescription
            Güven(1) = "1": Call Form_Load
        End If
        txtPass.SetFocus: txtPass.SelStart = 0: txtPass.SelLength = Len(txtPass.Text)
    End If
End Sub
Private Sub cmdCancel_Click()
    TheEnd
End Sub
Private Sub Form_Load()
    DenemeNo = 0
    If Güven(1) = "1" Then Label1.Caption = "Güvenlik kodunu giriniz": Caption = "Güvenlik Kodu" Else Label1.Caption = "Açma kodunu giriniz": Caption = "Açma Kodu"
    Write #7, Time$, Me.Name, "Form_Load", "Successful" 'logging
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Write #7, Time$, Me.Name, "Form_Unload", "Successful" 'logging
End Sub
