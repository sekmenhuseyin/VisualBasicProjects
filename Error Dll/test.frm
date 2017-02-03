VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hatalar"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3000
   Icon            =   "test.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Objext required"
      Height          =   540
      HelpContextID   =   5
      Left            =   105
      TabIndex        =   4
      Top             =   2625
      Width           =   2745
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Type mismatch"
      Height          =   540
      HelpContextID   =   4
      Left            =   105
      TabIndex        =   3
      Top             =   1995
      Width           =   2745
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Overflow"
      Height          =   540
      HelpContextID   =   3
      Left            =   128
      TabIndex        =   2
      Top             =   1365
      Width           =   2745
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Division by 0"
      Height          =   540
      HelpContextID   =   2
      Left            =   128
      TabIndex        =   1
      Top             =   735
      Width           =   2745
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dosya bulunamadý"
      Height          =   540
      HelpContextID   =   1
      Left            =   128
      TabIndex        =   0
      Top             =   105
      Width           =   2745
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Byte
Private Sub Command1_Click()
    On Error GoTo hataLog
    Open "temp.txt" For Input As #1
'''''''''''''''''''''''''''''''''''''
Exit Sub
hataLog:
    HataLogger
End Sub
Private Sub Command2_Click()
    On Error GoTo hataLog
    i = 5 / 0
'''''''''''''''''''''''''''''''''''''
Exit Sub
hataLog:
    HataLogger
End Sub
Private Sub Command3_Click()
    On Error GoTo hataLog
    i = 500000
'''''''''''''''''''''''''''''''''''''
Exit Sub
hataLog:
    HataLogger
End Sub
Private Sub Command4_Click()
    On Error GoTo hataLog
    i = "sekmen"
'''''''''''''''''''''''''''''''''''''
Exit Sub
hataLog:
    HataLogger
End Sub
Private Sub Command5_Click()
    On Error GoTo hataLog
    Load image1
'''''''''''''''''''''''''''''''''''''
Exit Sub
hataLog:
    HataLogger
End Sub
