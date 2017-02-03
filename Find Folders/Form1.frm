VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Klasör"
   ClientHeight    =   7110
   ClientLeft      =   4365
   ClientTop       =   2670
   ClientWidth     =   6630
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   6630
   Begin VB.CommandButton Command5 
      Caption         =   "CD-ROM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3360
      TabIndex        =   8
      Top             =   840
      Width           =   3000
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Bul"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   3000
   End
   Begin VB.ListBox List1 
      Height          =   4155
      ItemData        =   "Form1.frx":15162
      Left            =   240
      List            =   "Form1.frx":15164
      TabIndex        =   5
      Top             =   2640
      Width           =   6135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kullanýcý Adý"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3000
   End
   Begin VB.CommandButton Command_exit 
      Caption         =   "Çýkýþ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   6120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "System32"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Windows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3000
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   6825
      Left            =   105
      TabIndex        =   4
      Top             =   120
      Width           =   6360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetsystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const DRIVE_REMOVABLE = 2: Private Const DRIVE_FIXED = 3: Private Const DRIVE_REMOTE = 4: Private Const DRIVE_CDROM = 5: Private Const DRIVE_RAMDISK = 6
'''''''''''''''''''''''''''''''''''''''
Dim Yol As String: Dim Uzunluk As Integer: Dim i As Byte
Private Sub Command_exit_Click()
    End
End Sub
Private Sub Command1_Click()
    Yol = Space(255)
    Uzunluk = GetWindowsDirectory(Yol, Len(Yol))
    MsgBox Left(Yol, Uzunluk)
End Sub
Private Sub Command2_Click()
    Yol = Space(255)
    Uzunluk = GetsystemDirectory(Yol, Len(Yol))
    MsgBox Left(Yol, Uzunluk)
End Sub
Private Sub Command3_Click()
    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    MsgBox sBuffer
End Sub
Private Sub Command4_Click()
    If IsNumeric(Text1.Text) = True Then MsgBox Environ(Val(Text1.Text)) Else MsgBox Environ(Text1.Text)
End Sub
Private Sub Command5_Click()
    Dim r&, allDrives$, JustOneDrive$, pos%, DriveType&
    Dim CDfound As Integer
    allDrives$ = Space$(64)
    r& = GetLogicalDriveStrings(Len(allDrives$), allDrives$)
    allDrives$ = Left$(allDrives$, r&)
    Do
        pos% = InStr(allDrives$, Chr$(0))
        If pos% Then
            JustOneDrive$ = Left$(allDrives$, pos%)
            allDrives$ = Mid$(allDrives$, pos% + 1, Len(allDrives$))
            DriveType& = GetDriveType(JustOneDrive$)
            If DriveType& = DRIVE_CDROM Then
                CDfound% = True
                Exit Do
            End If
        End If
    Loop Until allDrives$ = "" Or DriveType& = DRIVE_CDROM
    If CDfound% Then
        MsgBox "The CD-ROM drive on your system is drive " & UCase$(JustOneDrive$)
    Else
        MsgBox "No CD-ROM drives were detected on your system."
    End If
End Sub
Private Sub Form_Load()
    For i = 1 To 40
        List1.AddItem Environ(i)
    Next i
End Sub
Private Sub List1_Click()
    Dim i As Integer: Text1.Text = ""
    For i = 1 To Len(List1.Text)
        If Mid(List1.List(List1.ListIndex), i, 1) = "=" Then Exit For
        Text1.Text = Text1.Text & Mid(List1.List(List1.ListIndex), i, 1)
    Next i
End Sub
