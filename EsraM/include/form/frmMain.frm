VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H00000000&
   ClientHeight    =   7545
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   4695
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'max menu remover
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'max buton remover
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Sub DisableMaxBtn()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''max buton remover
    Const WS_MAXIMIZEBOX = &H10000: Const GWL_STYLE = (-16): Dim L As Long
    L = GetWindowLong(Me.hwnd, GWL_STYLE)
    L = L And Not WS_MAXIMIZEBOX  'Disables Maximize Button
    L = SetWindowLong(Me.hwnd, GWL_STYLE, L)
End Sub
Private Sub MDIForm_Load()
    Dim hMenu As Long: Dim menuItemCount As Long: Dim theItem As Long: Dim i As Long: Const MF_BYPOSITION = &H400: Const MF_REMOVE = &H1000
    Write #7, Time$, Me.Name, "MDIForm_Load", "Start" 'logging
    Me.Show: Me.Move 0, 0: Me.Caption = App.ProductName ': DockingStart Me, True
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''max menu remover
    'Get system menu From Form         'Set initial variable For "not found"
    hMenu = GetSystemMenu(Me.hwnd, 0): theItem = -1
    If hMenu Then 'If found then
        menuItemCount = GetMenuItemCount(hMenu) 'Retrieve menu count
        For i = 0 To menuItemCount 'Look for the "max" bnuton
            If GetMenuItemID(hMenu, i) = 61488 Then
                theItem = i 'If found, quit Loop
                Exit For
            End If
        Next i
        If theItem <> -1 Then 'If found then remove Item
            Call RemoveMenu(hMenu, theItem, MF_REMOVE Or MF_BYPOSITION)
        End If
        Call DrawMenuBar(Me.hwnd) 'Draw it!
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''max menu remover
    Write #7, Time$, Me.Name, "MDIForm_Load", "Successful" 'logging
End Sub
Public Sub MDIForm_Resize()
    If Screen.ActiveForm.Name <> Me.Name Then
        Me.Width = Screen.ActiveForm.Width + 500: Me.Height = Screen.ActiveForm.Height + 1000
        Screen.ActiveForm.Top = (Me.ScaleHeight - Screen.ActiveForm.Height) / 2
        Screen.ActiveForm.Left = (Me.ScaleWidth - Screen.ActiveForm.Width) / 2
    End If
    Call DisableMaxBtn
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    'DockingTerminate Me
    Write #7, Time$, Me.Name, "MDIForm_Unload", "Successful" 'logging
End Sub
