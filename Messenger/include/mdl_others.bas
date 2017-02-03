Attribute VB_Name = "mdl_others"
Option Explicit
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
'AlwaysOnTop
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const SWP_NOACTIVATE = &H10: Private Const SWP_SHOWWINDOW = &H40: Private Const HWND_NOTOPMOST = -2: Private Const HWND_TOPMOST = -1
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
'MakeTransparent
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+

Public Sub AlwaysOnTop(FormName As Form, SetOnTop As Boolean)
    Dim lFlag
    If SetOnTop Then lFlag = HWND_TOPMOST Else lFlag = HWND_NOTOPMOST
    SetWindowPos FormName.hwnd, lFlag, _
    FormName.Left / Screen.TwipsPerPixelX, FormName.Top / Screen.TwipsPerPixelY, _
    FormName.Width / Screen.TwipsPerPixelX, FormName.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub
Public Sub MakeTransparent(HandleNo As Long, AlphaVal As Byte)
    SetWindowLong HandleNo, GWL_EXSTYLE, GetWindowLong(HandleNo, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes HandleNo, vbBlack, AlphaVal, &H2
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Public Function UpperCaseFirstLetter(ByVal OldString As String) As String
    If Trim(OldString) = "" Then Exit Function: Dim i As Integer
    'ilk önce ilk harf hariç tüm harfleri küçültüyoruz.
    For i = 2 To Len(OldString)
        If Mid(OldString, i, 1) = "Ý" Then
            Mid(OldString, i, 1) = "i"
        ElseIf Mid(OldString, i, 1) = "I" Then
            Mid(OldString, i, 1) = "ý"
        Else
            Mid(OldString, i, 1) = LCase(Mid(OldString, i, 1))
        End If
    Next i
    i = 0
basadön:
    i = i + 1
    'daha sonra sadece ilk harfi büyütüyoruz.
    If Mid(OldString, i, 1) = "i" Then
        Mid(OldString, i, 1) = "Ý"
    ElseIf Mid(OldString, i, 1) = "ý" Then
        Mid(OldString, i, 1) = "I"
    Else
        Mid(OldString, i, 1) = UCase(Mid(OldString, i, 1))
    End If
    'baþka kelime varsa onlarýn da ilk harflerini büyütecez
    i = i + 1: i = InStr(i, OldString, " ")
    If i <> 0 Then GoTo basadön
    'en son olarak da yukarýda bulduðumuz iki sonucu birleþtirip geriye o stringi döndürüyoruz.
    UpperCaseFirstLetter = OldString
End Function
Public Function RenkTemalarý(Theme_Name As String, index As Byte) As String
    Select Case Theme_Name
    Case "Standart"
        If index = 0 Then
                               RenkTemalarý = RGB(255, 192, 128) 'arka renk
        ElseIf index = 1 Then: RenkTemalarý = RGB(255, 224, 192) 'yazý alaný rengi
        ElseIf index = 2 Then: RenkTemalarý = RGB(0, 0, 0)       'yazý rengi
        ElseIf index = 3 Then: RenkTemalarý = RGB(255, 192, 128) 'buton arka rengi
        ElseIf index = 4 Then: RenkTemalarý = RGB(0, 0, 0)       'buton yazý rengi
        End If
    Case "Mavimsi"
        If index = 0 Then
                               RenkTemalarý = RGB(0, 128, 192)
        ElseIf index = 1 Then: RenkTemalarý = RGB(75, 180, 242)
        ElseIf index = 2 Then: RenkTemalarý = RGB(0, 0, 0)
        ElseIf index = 3 Then: RenkTemalarý = RGB(50, 150, 200)
        ElseIf index = 4 Then: RenkTemalarý = RGB(0, 0, 0)
        End If
    Case "Kömür Karasý"
        If index = 0 Then
                               RenkTemalarý = RGB(0, 0, 0)
        ElseIf index = 1 Then: RenkTemalarý = RGB(64, 64, 64)
        ElseIf index = 2 Then: RenkTemalarý = RGB(255, 255, 255)
        ElseIf index = 3 Then: RenkTemalarý = RGB(64, 64, 64)
        ElseIf index = 4 Then: RenkTemalarý = RGB(255, 255, 255)
        End If
    Case "Windows XP"
        If index = 0 Then
                               RenkTemalarý = -2147483633
        ElseIf index = 1 Then: RenkTemalarý = -2147483643
        ElseIf index = 2 Then: RenkTemalarý = -2147483640
        ElseIf index = 3 Then: RenkTemalarý = -2147483633
        ElseIf index = 4 Then: RenkTemalarý = -2147483640
        End If
    End Select
End Function
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+

