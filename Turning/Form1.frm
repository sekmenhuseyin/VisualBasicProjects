VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2040
      Top             =   4320
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1440
      Shape           =   2  'Oval
      Top             =   2040
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i
Private Sub Timer1_Timer()
    If i < 360 Then i = i + 1 Else i = 0
    Shape1.Left = Sin(i) * 2000 + 4000
    Shape1.Top = Cos(i) * 2000 + 3000
End Sub
