VERSION 5.00
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form frm_Change 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deðiþtir"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   Icon            =   "frm_Change.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   60
      Top             =   3060
   End
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   3645
      TabIndex        =   4
      Top             =   1965
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
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
      MICON           =   "frm_Change.frx":15162
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
      Height          =   435
      Left            =   165
      TabIndex        =   3
      Top             =   1965
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
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
      MICON           =   "frm_Change.frx":1517E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtNew2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2205
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1428
      Width           =   3375
   End
   Begin VB.TextBox txtNew1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2205
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   888
      Width           =   3375
   End
   Begin VB.TextBox txtOld 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2205
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   161
      Width           =   3375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Yeni Kodunuz (Tekrar)"
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
      Left            =   165
      TabIndex        =   7
      Top             =   1518
      Width           =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Yeni Kodunuz"
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
      Left            =   900
      TabIndex        =   6
      Top             =   978
      Width           =   1185
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
      Left            =   915
      TabIndex        =   5
      Top             =   251
      Width           =   1170
   End
End
Attribute VB_Name = "frm_Change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Write #7, Time$, Me.Name, "cmdOK_Click", "Start" 'logging
    Dim DeðiþecekKod, CodeType As String
    If Me.Caption = "Açma Kodunu Deðiþtir" Then DeðiþecekKod = Güven(4): CodeType = "açma" Else DeðiþecekKod = Güven(5): CodeType = "güvenlik"
    If CalculateMD5(txtOld.Text) = DeðiþecekKod Then
        If txtNew1.Text = txtNew2.Text Then
            If txtNew1.Text = "" Then
                MsgBox UpperCaseFirstLetter(CodeType) & " kodu boþ olamaz !", , App.FileDescription
                txtNew1.SetFocus: txtNew1.SelStart = 0: txtNew1.SelLength = Len(txtNew1.Text)
                Write #7, Time$, Me.Name, "cmdOK_Click", "NewPassEmpty" 'logging
            Else
                If CodeType = "açma" Then Güven(4) = CalculateMD5(txtNew1.Text) Else Güven(5) = CalculateMD5(txtNew1.Text)
                MsgBox UpperCaseFirstLetter(CodeType) & " kodunuz baþarýyla deðiþtirilmiþtir !", , App.FileDescription
                Write #7, Time$, Me.Name, "cmdOK_Click", "SuccessfulPassChange" 'logging
                Unload Me
            End If
        Else
            MsgBox "Yeni " & CodeType & " kodlarýnýz birbiriyle uyuþmuyor !", , App.FileDescription
            txtNew1.SetFocus: txtNew1.SelStart = 0: txtNew1.SelLength = Len(txtNew1.Text)
            Write #7, Time$, Me.Name, "cmdOK_Click", "NewPassWrong" 'logging
        End If
    Else
        MsgBox "Eski " & CodeType & " kodunuz hatalý !", , App.FileDescription
        txtOld.SetFocus: txtOld.SelStart = 0: txtOld.SelLength = Len(txtOld.Text)
        Write #7, Time$, Me.Name, "cmdOK_Click", "OldPassWrong" 'logging
    End If
    Write #7, Time$, Me.Name, "cmdOK_Click", "End" 'logging
End Sub
Private Sub Form_Load()
    If Ayar(5) = "1" Then AlwaysOnTop Me, True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frm_Options.Enabled = True: frm_Options.Güven0.SetFocus
    Write #7, Time$, Me.Name, "Form_Unload", "Successful" 'logging
End Sub
Private Sub Timer1_Timer()
    Write #7, Time$, Me.Name, "Timer1_Timer", Me.Caption 'logging
    If Me.Caption = "0" Then 'açma kodunu deðiþtir
        Me.Caption = "Açma Kodunu Deðiþtir"
    ElseIf Me.Caption = "1" Then 'güvenlik kodunu deðiþtir
        Me.Caption = "Güvenlik Kodunu Deðiþtir"
    End If
    Write #7, Time$, Me.Name, "Timer1_Timer", "End" 'logging
    Timer1.Enabled = False
End Sub
