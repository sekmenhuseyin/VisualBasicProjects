VERSION 5.00
Object = "{9898A615-558F-4B19-8D75-2D993AD12970}#1.0#0"; "TrayIcon.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin TrayIconOCX.ToolTipOnDemand ToolTipOnDemand1 
      Left            =   1890
      Top             =   1050
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin TrayIconOCX.TrayIcon TrayIcon1 
      Left            =   1575
      Top             =   1050
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.Menu mn_Tray 
      Caption         =   "tray"
      Visible         =   0   'False
      Begin VB.Menu mn_Tray_Show 
         Caption         =   "göster"
      End
      Begin VB.Menu mN_Tray_Hide 
         Caption         =   "sakla"
      End
      Begin VB.Menu mn_Tray_Exit 
         Caption         =   "çýkýþ"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Me.Caption = App.Title
    With TrayIcon1
        .IconHandle = Me.Icon
        .ToolTip = App.Title
        .Create Me.hWnd
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    TrayIcon1.Remove
    ToolTipOnDemand1.Destroy
End Sub
Private Sub mn_Tray_Exit_Click()
    Unload Me
End Sub
Private Sub mN_Tray_Hide_Click()
    Me.Hide
    Call ShowBalloon(TrayIcon1.SysTrayHWnd, btInfo, "I'm here", Me.Caption)
End Sub
Private Sub mn_Tray_Show_Click()
    Me.Show: Me.SetFocus
End Sub
Private Sub ToolTipOnDemand1_BalloonDestroyed()
    TrayIcon1.TrackIconMovement = False
End Sub
Private Sub ToolTipOnDemand1_BalloonShowed()
    TrayIcon1.TrackIconMovement = True
End Sub
Private Sub TrayIcon1_TrayMouseEvent(ByVal MouseEvent As stMouseEvent)
    Select Case MouseEvent
        Case stMouseMove
        Case stLeftButtonDown
            mn_Tray.Visible = False
        Case stLeftButtonUp
            Call mn_Tray_Show_Click
        Case stLeftButtonDoubleClick
        Case stRightButtonDown
            mn_Tray.Visible = False
        Case stRightButtonUp
            PopupMenu mn_Tray
        Case stRightButtonDoubleClick
        Case stMiddleButtonDown
        Case stMiddleButtonUp
        Case stMiddleButtonDoubleClick
    End Select
End Sub
Private Sub ToolTipOnDemand1_MouseEvents(MouseEvent As Long)
    Select Case MouseEvent
        Case stMouseMove
        Case stLeftButtonDown
            ToolTipOnDemand1.Destroy
        Case stLeftButtonUp
        Case stLeftButtonDoubleClick
        Case stRightButtonDown
            ToolTipOnDemand1.Destroy
        Case stRightButtonUp
        Case stRightButtonDoubleClick
        Case stMiddleButtonDown
        Case stMiddleButtonUp
        Case stMiddleButtonDoubleClick
    End Select
End Sub
Private Sub ShowBalloon(ByVal SystemTrayIndex As Long, ByVal enIconType As blIconType, ByVal sPrompt As String, Optional ByVal sTitle As String, _
                        Optional ByVal lTimeout As Long = 5000, Optional ByVal lBackColor As Long = -1, Optional ByVal lForeColor As Long = -1)
    On Local Error Resume Next
    Dim lX As Long, lY As Long
    Call TrayIcon1.GetIconMiddle(lX, lY)
    If lForeColor = -1 Then lForeColor = vbBlack
    If lBackColor = -1 Then lBackColor = &H80000018
    With ToolTipOnDemand1
        .ParentHwnd = SystemTrayIndex
        .x = lX
        .y = lY
        .BackColor = lBackColor
        .ForeColor = lForeColor
        .Prompt = sPrompt
        .Title = sTitle
        .TimeOut = lTimeout
        .IconType = enIconType
        .Show
      End With
End Sub

