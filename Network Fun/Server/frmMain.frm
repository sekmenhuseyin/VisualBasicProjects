VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client PC Controller"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5685
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   1111
      ButtonWidth     =   1561
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "File"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Find Computers"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Exit"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Animations"
            Style           =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Style"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   21
      Top             =   6765
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List2 
      Height          =   450
      ItemData        =   "frmMain.frx":15162
      Left            =   210
      List            =   "frmMain.frx":15164
      TabIndex        =   20
      Top             =   2730
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Frame Frame3 
      Caption         =   "Client Operations"
      Height          =   1590
      Left            =   210
      TabIndex        =   13
      Top             =   3465
      Width           =   5265
      Begin VB.CommandButton cmdClientKill 
         Caption         =   "Kill"
         Height          =   495
         Left            =   1485
         TabIndex        =   18
         Top             =   945
         Width           =   1335
      End
      Begin VB.CommandButton cmd_Unlock 
         Caption         =   "Unlock Screens"
         Height          =   495
         Left            =   1485
         TabIndex        =   17
         Top             =   315
         Width           =   1335
      End
      Begin VB.CommandButton cmd_Lock 
         Caption         =   "Lock Screens"
         Height          =   495
         Left            =   105
         TabIndex        =   16
         Top             =   315
         Width           =   1335
      End
      Begin VB.CommandButton cmdClientUpdate 
         Caption         =   "Upgrade Client"
         Height          =   495
         Left            =   2865
         TabIndex        =   15
         Top             =   945
         Width           =   1335
      End
      Begin VB.CommandButton cmdClientRun 
         Caption         =   "Run"
         Height          =   495
         Left            =   105
         TabIndex        =   14
         Top             =   945
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Animations"
      Height          =   1380
      Left            =   198
      TabIndex        =   7
      Top             =   5145
      Width           =   5295
      Begin VB.Timer TimerAnimation 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   4830
         Top             =   105
      End
      Begin VB.CommandButton cmdEndAni 
         Caption         =   "End Show"
         Height          =   810
         Left            =   4095
         TabIndex        =   19
         Top             =   375
         Width           =   1005
      End
      Begin VB.CommandButton cmdShowAni 
         Caption         =   "Start Show"
         Height          =   810
         Left            =   3015
         TabIndex        =   9
         Top             =   375
         Width           =   1005
      End
      Begin VB.ComboBox theAniType 
         Height          =   315
         ItemData        =   "frmMain.frx":15166
         Left            =   750
         List            =   "frmMain.frx":15170
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   375
         Width           =   2160
      End
      Begin MSComctlLib.Slider theAniSpeed 
         Height          =   255
         Left            =   750
         TabIndex        =   10
         Top             =   870
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   20
         SelStart        =   10
         TickFrequency   =   2
         Value           =   10
      End
      Begin VB.Label Label2 
         Caption         =   "Type"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   375
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   870
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Font Options"
      Height          =   1095
      Left            =   2471
      TabIndex        =   4
      Top             =   765
      Width           =   3015
      Begin VB.CommandButton cmdSelectColor 
         Caption         =   "Background"
         Height          =   495
         Left            =   1785
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSelectFont 
         Caption         =   "Font"
         Height          =   495
         Left            =   225
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   1305
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
   End
   Begin VB.CommandButton cmdListPC 
      Caption         =   "List Computers"
      Height          =   495
      Left            =   191
      TabIndex        =   2
      Top             =   795
      Width           =   2025
   End
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "frmMain.frx":151A0
      Left            =   195
      List            =   "frmMain.frx":151A2
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   1380
      Width           =   2025
   End
   Begin VB.CommandButton cmdSendMsg 
      Caption         =   "Send Message"
      Default         =   -1  'True
      Height          =   540
      Left            =   2471
      TabIndex        =   1
      Top             =   2685
      Width           =   3015
   End
   Begin VB.TextBox txt_MsgText 
      Height          =   375
      Left            =   2471
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2160
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim T_Font, T_Size, T_Color, T_Bold, T_Italic, T_Strike, T_Underline, T_Back, CPath As String: Dim i, j As Integer
Dim k, l, a, b, c, d, e As Integer
Private Sub cmd_Lock_Click(): Call CopySelected: Call Writingprocess("order", "lockscreen"): End Sub
Private Sub cmd_Unlock_Click(): Call CopySelected: Call Writingprocess("order", "unlockscreen"): End Sub
Private Sub cmdClientKill_Click(): Call CopySelected: Call Writingprocess("order", "die"): End Sub
Private Sub cmdSendMsg_Click(): Call CopySelected: Call Writingprocess("write", txt_MsgText.Text): End Sub
Private Sub cmdShowAni_Click(): Call CopySelected: c = 1: TimerAnimation.Enabled = True: i = 0: Call TimerAnimation_Timer: End Sub
Private Sub cmdEndAni_Click(): TimerAnimation.Enabled = False: End Sub
Private Sub theAniSpeed_Change(): TimerAnimation.Interval = Int(5000 / theAniSpeed): End Sub
Private Sub TimerAnimation_Timer()
    j = j + 1: k = List2.ListCount: l = Len(txt_MsgText.Text)
    Select Case theAniType.ListIndex
        Case 0 'tum bilgisayarlara mesaji harf harf gonderir
            Writingprocess "write", Mid(txt_MsgText, j, 1)
        Case 1 'bilgisayarlarda kayan yazi misali harfler kayarak belirir  --  insallah
            'ilk dongu bilgisayar sayisi dongusu. ilk basta 0 sonrakinda 1,2,3 vs...
            
            
            'merak ettigim 5 pc ve 1 har oldugunda ne oluyor
            've 5 harf 1 2 pc oldugunda ne oluyor
            'yada ne olmali
            For a = 0 To c - 1
                d = c - a: If d > l Then d = d - l
                e = a: If e = k Then e = e - k
                'ilk dongude 0.pcye 1.harfi gonderiyor
                'ikinci dongude 0.pcye 2.harf, 1.pcye 1.harf geliyor (asagide c++ oluyor)
                Writingprocess "write", Mid(txt_MsgText, c - a, 1), Val(e)
            Next a
            c = c + 1: If c > l Then c = 0
        Case 2 'yanip sonen isik misali yanip sonen text1
    End Select
    If j = Len(txt_MsgText.Text) Then j = 0
End Sub
Private Sub cmdClientRun_Click()
    On Error GoTo TheErr
    'loop kullan selected pc's icin
    Call CopySelected
    For i = 0 To List2.ListCount - 1
        Shell "\\" & List2.List(i) & "\c$\nc.exe", vbNormalFocus
    Next i
    Exit Sub
TheErr:
    MsgBox "Couldn't find the client !", vbExclamation
End Sub
Private Sub cmdClientUpdate_Click()
    On Error GoTo TheErr
    Call cmdClientKill_Click
    With CD1
        .DialogTitle = "Choose Client Software"
        .Flags = cdlOFNFileMustExist
        .InitDir = CPath
        .FileName = "nc.exe"
        .Filter = "Client Software|nc.exe"
        .ShowOpen
        CPath = .FileName
    End With
    'now the erase the client software from the selected pc's and copy the new version
    'use a loop for selected pc's
    Call CopySelected
    For i = 0 To List2.ListCount - 1
        Kill "\\" & List2.List(i) & "\c$\nc.exe": FileCopy CPath, "\\" & List2.List(i) & "\c$\nc.exe"
    Next i
    Exit Sub
TheErr:
End Sub
Private Sub cmdListPC_Click() 'find the pc's on the network
    Dim l As New cm_LAN: Dim s() As String
    Me.MousePointer = 11: List1.Clear: List1.AddItem "Working...": s = Split(l.GetPCList, "||"): List1.Clear
    For i = LBound(s) To UBound(s)
        If s(i) <> MyName Then List1.AddItem s(i)
    Next i
    List1.ListIndex = 0: Me.MousePointer = 1
End Sub
Private Sub cmdSelectColor_Click() 'select background color
    On Error GoTo TheErr
    CD1.Color = T_Back: CD1.ShowColor: T_Back = CD1.Color
    Exit Sub
TheErr:
End Sub
Private Sub cmdSelectFont_Click() 'select font
    On Error GoTo TheErr
    With CD1
        .DialogTitle = "Font"   'set the dialog title
        .Flags = cdlCFBoth Or cdlCFEffects   'set the flags so you can access the fonts* and the strikethru, underline, color
        .Color = T_Color: .FontName = T_Font: .FontSize = T_Size: .FontBold = T_Bold: .FontItalic = T_Italic: .FontStrikethru = T_Strike: .FontUnderline = T_Underline
        .ShowFont  'show the dialog     'now replace the givens with our variables
        T_Font = .FontName: T_Size = .FontSize: T_Color = .Color: T_Bold = .FontBold: T_Italic = .FontItalic: T_Strike = .FontStrikethru: T_Underline = .FontUnderline
    End With
    Exit Sub
TheErr:
End Sub
Private Sub Form_Load()
    T_Font = "Arial Black": T_Size = 250: T_Color = 16777215: T_Bold = True: T_Italic = False: T_Strike = False: T_Underline = False: T_Back = 0
    CPath = App.Path:: theAniType.ListIndex = 0
    List1.AddItem MyName
    List1.AddItem "STUDENT24"
    List1.AddItem "STUDENT23"
    List1.AddItem "STUDENT22"
    List1.AddItem "STUDENT21"
    List1.AddItem "STUDENT20"
    List1.ListIndex = 0 ': List1.Selected(0) = True
End Sub
Private Sub Writingprocess(Order As String, MsgText As String, Optional pcID As Integer)
    If Val(pcID) > 0 Then
        Writingprocess2 Order, MsgText, pcID
    Else
        For i = 0 To List2.ListCount - 1
            Writingprocess2 Order, MsgText, Val(i)
        Next i
    End If
End Sub
Private Sub Writingprocess2(Order As String, MsgText As String, Optional pcID As Integer)
    Open "\\" & List2.List(pcID) & "\c$\network.txt" For Output As #1
    Write #1, Order, MsgText, T_Font, T_Size, T_Color, T_Bold, T_Italic, T_Strike, T_Underline, T_Back
    Close #1
End Sub
Private Sub CopySelected()
    List2.Clear
    If List1.SelCount = 0 Then
        List2.AddItem List1.List(List1.ListIndex)
    Else
        For i = 0 To List1.ListCount - 1
            If List1.Selected(i) = True Then List2.AddItem List1.List(i)
        Next i
    End If
End Sub
