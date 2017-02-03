VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   2325
   ClientTop       =   915
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9120
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   420
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Çýkýþ"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   1050
      TabIndex        =   0
      Top             =   1050
      Width           =   2325
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   9135
      X2              =   8925
      Y1              =   6930
      Y2              =   7140
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   9135
      X2              =   8820
      Y1              =   6825
      Y2              =   7140
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   8820
      MousePointer    =   8  'Size NW SE
      TabIndex        =   1
      Top             =   6825
      Width           =   330
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SagSol As Single: Dim AsagýYukarý As Single: Dim IlkSagSol As Single: Dim IlkAsagýYukarý As Single
Dim IlkTýkX As Single: Dim IlkTýkY As Single: Dim TýkX As Single: Dim TýkY As Single
Dim OrgX As Single: Dim OrgY As Single
Private Sub Command1_Click()
    End
End Sub
Private Sub Form_Load()
    Label1.Left = Form1.Width - Label1.Width
    Label1.Top = Form1.Height - Label1.Height
    Line1.Y1 = Label1.Top: Line1.X1 = Label1.Left + Label1.Width
    Line1.Y2 = Label1.Top + Label1.Height: Line1.X2 = Label1.Left
    Line2.Y1 = Label1.Top: Line2.X1 = Label1.Left + Label1.Width + 100
    Line2.Y2 = Label1.Top + Label1.Height: Line2.X2 = Label1.Left + 100
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    IlkSagSol = X
    IlkAsagýYukarý = Y
    Form1.MousePointer = 15
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SagSol = X
    AsagýYukarý = Y
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False
    Form1.MousePointer = 0
End Sub
Private Sub Form_Resize()
    If Me.Width < 5000 Then Me.Width = 5000
    If Me.Height < 4000 Then Me.Height = 4000
End Sub
Private Sub Timer1_Timer()
    OrgX = Me.Left + SagSol - IlkSagSol: OrgY = Me.Top + AsagýYukarý - IlkAsagýYukarý
    If OrgX < 250 And OrgX > -250 Then
        Me.Left = 0
    ElseIf OrgX < Screen.Width - Me.Width + 250 And OrgX > Screen.Width - Me.Width - 100 Then: Me.Left = Screen.Width - Me.Width
    Else: Me.Left = Me.Left + SagSol - IlkSagSol: End If
    If OrgY < 250 And OrgY > -250 Then
        Me.Top = 0
    ElseIf OrgY < Screen.Height - Me.Height + 250 And OrgY > Screen.Height - Me.Height - 100 Then: Me.Top = Screen.Height - Me.Height
    Else: Me.Top = Me.Top + AsagýYukarý - IlkAsagýYukarý: End If
'    Form1.Left = Form1.Left + SagSol - IlkSagSol
'    Form1.Top = Form1.Top + AsagýYukarý - IlkAsagýYukarý
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer2.Enabled = True
    IlkTýkX = X: IlkTýkY = Y
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TýkX = X
    TýkY = Y
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer2.Enabled = False
End Sub
Private Sub Timer2_Timer()
    On Error Resume Next
    Form1.Width = Form1.Width + TýkX - IlkTýkX
    Form1.Height = Form1.Height + TýkY - IlkTýkY
    Label1.Top = Form1.Height - Label1.Height
    Label1.Left = Form1.Width - Label1.Width
    Line1.Y1 = Label1.Top: Line1.X1 = Label1.Left + Label1.Width
    Line1.Y2 = Label1.Top + Label1.Height: Line1.X2 = Label1.Left
    Line2.Y1 = Label1.Top: Line2.X1 = Label1.Left + Label1.Width + 100
    Line2.Y2 = Label1.Top + Label1.Height: Line2.X2 = Label1.Left + 100
End Sub
