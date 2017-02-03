VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5520
      Top             =   4680
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   360
      Shape           =   2  'Oval
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim toLeft, toTop As Integer

Private Sub Form_Load()
    toLeft = 30
    toTop = 30
End Sub

Private Sub Timer1_Timer()
    Shape1.Left = Shape1.Left + toLeft
    Shape1.Top = Shape1.Top + toTop
    If Shape1.Left < 0 Or Shape1.Left > Form1.Width - Shape1.Width Then toLeft = -toLeft
    If Shape1.Top < 0 Or Shape1.Top > Form1.Height - Shape1.Height - Shape1.Height Then toTop = -toTop
End Sub
