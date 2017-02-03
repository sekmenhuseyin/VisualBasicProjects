VERSION 5.00
Begin VB.Form Kimlik 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kimlik Listesi"
   ClientHeight    =   5520
   ClientLeft      =   630
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "Kimlik.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   -1  'True
      Height          =   345
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "veriler"
      Top             =   4560
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Anasayfa"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   408
      TabIndex        =   0
      Top             =   4615
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      ItemData        =   "Kimlik.frx":27A2
      Left            =   405
      List            =   "Kimlik.frx":27A4
      TabIndex        =   1
      Top             =   1050
      Width           =   7995
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sýra   Adý                  Doðum         Sicil              No     Soyadý               Tarihi        No"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   420
      TabIndex        =   3
      Top             =   525
      Width           =   7995
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Height          =   5100
      Left            =   203
      TabIndex        =   2
      Top             =   210
      Width           =   8505
   End
End
Attribute VB_Name = "Kimlik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim no, ad, dog As String
Private Sub Command1_Click()
On Error Resume Next
Data1.Recordset.Close
Giriþ.Show: Unload Me
End Sub
Private Sub Form_Activate()
On Error GoTo son
List1.Clear
With Data1.Recordset
.MoveFirst
bas:
List1.AddItem .Fields("No") & String(5 - Len(.Fields("no")), " ") & .Fields("Ad Soyad") & String(22 - Len(.Fields("Ad Soyad")), " ") & .Fields("Doðum Tarihi") & String(13 - Len(.Fields("Doðum Tarihi")), " ") & .Fields("Sicil No")
'List1.AddItem .Fields("No") & String(5 - .Fields("No"), " ") & .Fields("Ad Soyad") & String(40 - .Fields("Ad Soyad"), " ") & .Fields("Doðum Tarihi") & String(17 - .Fields("Doðum Tarihi"), " ") & .Fields("Sicil No")
.MoveNext
If .EOF = True Then GoTo son
GoTo bas
son:
End With
End Sub
Private Sub Form_Load()
Data1.DatabaseName = App.Path + "\3308.mdb"
End Sub
Private Sub Form_Unload(Cancel As Integer)
Command1_Click
End Sub
