VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form f_Main 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shot - Menu"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6900
   Icon            =   "frm_Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin MSForms.CommandButton cmd_Exit 
      Cancel          =   -1  'True
      Height          =   855
      Left            =   1328
      TabIndex        =   3
      Top             =   3255
      Width           =   4245
      Caption         =   "Çýkýþ"
      Size            =   "7488;1508"
      FontName        =   "Arial Black"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   162
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmd_About 
      Height          =   855
      Left            =   1328
      TabIndex        =   2
      Top             =   2415
      Width           =   4245
      Caption         =   "Hakkýnda"
      Size            =   "7488;1508"
      FontName        =   "Arial Black"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   162
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmd_Bests 
      Height          =   855
      Left            =   1328
      TabIndex        =   1
      Top             =   1575
      Width           =   4245
      Caption         =   "Yüksek Puanlar"
      Size            =   "7488;1508"
      FontName        =   "Arial Black"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   162
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmd_New 
      Default         =   -1  'True
      Height          =   855
      Left            =   1328
      TabIndex        =   0
      Top             =   735
      Width           =   4245
      Caption         =   "Yeni Oyun"
      Size            =   "7488;1508"
      FontName        =   "Arial Black"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   162
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "f_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_New_Click()
    f_Game.Top = Me.Top: f_Game.Left = Me.Left: f_Game.Show: Me.Hide
End Sub
Private Sub cmd_Bests_Click()
    f_Scores.Top = Me.Top: f_Scores.Left = Me.Left: f_Scores.Show: Me.Hide
End Sub
Private Sub cmd_Exit_Click()
    The_End
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 27: Unload Me: End
    End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    The_End
End Sub
