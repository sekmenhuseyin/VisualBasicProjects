VERSION 5.00
Begin VB.Form frm_Error 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HATA"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   ClipControls    =   0   'False
   Icon            =   "frm_Err.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDetails 
      Caption         =   "&Ayrýntýlarý Sakla <<"
      Height          =   375
      Left            =   4145
      TabIndex        =   3
      Top             =   1305
      Width           =   2085
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Tamam"
      Default         =   -1  'True
      Height          =   375
      Left            =   2885
      TabIndex        =   2
      Top             =   1305
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   1300
      Left            =   175
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2000
      Width           =   6040
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Eðer bu hata devam ederse uygulamanýzýn üreticisiyle iletiþim kurun."
      Height          =   195
      Left            =   945
      TabIndex        =   4
      Top             =   735
      Width           =   4755
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   175
      X2              =   6215
      Y1              =   1820
      Y2              =   1820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bir hata oluþtu."
      Height          =   195
      Left            =   945
      TabIndex        =   0
      Top             =   360
      Width           =   1050
   End
   Begin VB.Image imgError 
      Height          =   480
      Left            =   175
      Picture         =   "frm_Err.frx":5D52
      Top             =   200
      Width           =   480
   End
End
Attribute VB_Name = "frm_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDetails_Click()
    If cmdDetails.Caption = "&Ayrýntýlarý Sakla <<" Then
        cmdDetails.Caption = "&Ayrýntýlarý Göster >>"
        Me.Height = 2250
    Else
        cmdDetails.Caption = "&Ayrýntýlarý Sakla <<"
        Me.Height = 4000
    End If
End Sub
Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    If Dir(App.Path + "\report.err") = "" Then Unload Me: End
    Dim Number, Description, Source, HelpContext, HelpFile, LastDllError As String
    Open App.Path + "\report.err" For Input As #1
        Input #1, Number
        Input #1, Description
        Input #1, Source
        Input #1, HelpContext
        Input #1, HelpFile
        Input #1, LastDllError
    Close #1
    Kill App.Path + "\report.err"
    Call cmdDetails_Click
    Text1.Text = "Error Details:" & vbCrLf & vbCrLf & _
                 "Number: " & Number & vbCrLf & _
                 "Description: " & Description & vbCrLf & _
                 "Source: " & Source & vbCrLf & _
                 "HelpContext: " & HelpContext & vbCrLf & _
                 "HelpFile: " & HelpFile & vbCrLf & _
                 "LastDllError: " & LastDllError
End Sub
