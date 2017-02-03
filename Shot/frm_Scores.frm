VERSION 5.00
Begin VB.Form f_Scores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shot - Best Scores"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   Icon            =   "frm_Scores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "f_Scores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    f_Main.Top = Me.Top: f_Main.Left = Me.Left: f_Main.Show
End Sub

