Attribute VB_Name = "mdTray"
Option Explicit
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    UID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
'// constantes para la api Shell_NotifyIcon
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
'// constantes para capturar los eventos del formulario
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub ColocarIcono(ByVal hwnd As Long, ByVal hIcon As Long, ByVal sToolTip As String, SYStrayTipi As String)
Dim udtNOTIFYICONDATA As NOTIFYICONDATA
With udtNOTIFYICONDATA
   .cbSize = Len(udtNOTIFYICONDATA)
   .hwnd = hwnd
   .UID = vbNull   '.UID = 1&
   .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   .uCallbackMessage = WM_MOUSEMOVE
   .hIcon = hIcon
   If IsEmpty(sToolTip) Then
      .szTip = "" & vbNullChar
   Else
      .szTip = sToolTip & vbNullChar
   End If
End With

'agregar a la Barra
Select Case SYStrayTipi
    Case "ekle"
        Shell_NotifyIcon NIM_ADD, udtNOTIFYICONDATA
    Case "sil"
        Shell_NotifyIcon NIM_DELETE, udtNOTIFYICONDATA
    Case "deðiþtir"
        Shell_NotifyIcon NIM_MODIFY, udtNOTIFYICONDATA
End Select
End Sub
