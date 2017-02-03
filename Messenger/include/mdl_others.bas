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
    'ilk �nce ilk harf hari� t�m harfleri k���lt�yoruz.
    For i = 2 To Len(OldString)
        If Mid(OldString, i, 1) = "�" Then
            Mid(OldString, i, 1) = "i"
        ElseIf Mid(OldString, i, 1) = "I" Then
            Mid(OldString, i, 1) = "�"
        Else
            Mid(OldString, i, 1) = LCase(Mid(OldString, i, 1))
        End If
    Next i
    i = 0
basad�n:
    i = i + 1
    'daha sonra sadece ilk harfi b�y�t�yoruz.
    If Mid(OldString, i, 1) = "i" Then
        Mid(OldString, i, 1) = "�"
    ElseIf Mid(OldString, i, 1) = "�" Then
        Mid(OldString, i, 1) = "I"
    Else
        Mid(OldString, i, 1) = UCase(Mid(OldString, i, 1))
    End If
    'ba�ka kelime varsa onlar�n da ilk harflerini b�y�tecez
    i = i + 1: i = InStr(i, OldString, " ")
    If i <> 0 Then GoTo basad�n
    'en son olarak da yukar�da buldu�umuz iki sonucu birle�tirip geriye o stringi d�nd�r�yoruz.
    UpperCaseFirstLetter = OldString
End Function
Public Function RenkTemalar�(Theme_Name As String, index As Byte) As String
    Select Case Theme_Name
    Case "Standart"
        If index = 0 Then
                               RenkTemalar� = RGB(255, 192, 128) 'arka renk
        ElseIf index = 1 Then: RenkTemalar� = RGB(255, 224, 192) 'yaz� alan� rengi
        ElseIf index = 2 Then: RenkTemalar� = RGB(0, 0, 0)       'yaz� rengi
        ElseIf index = 3 Then: RenkTemalar� = RGB(255, 192, 128) 'buton arka rengi
        ElseIf index = 4 Then: RenkTemalar� = RGB(0, 0, 0)       'buton yaz� rengi
        End If
    Case "Mavimsi"
        If index = 0 Then
                               RenkTemalar� = RGB(0, 128, 192)
        ElseIf index = 1 Then: RenkTemalar� = RGB(75, 180, 242)
        ElseIf index = 2 Then: RenkTemalar� = RGB(0, 0, 0)
        ElseIf index = 3 Then: RenkTemalar� = RGB(50, 150, 200)
        ElseIf index = 4 Then: RenkTemalar� = RGB(0, 0, 0)
        End If
    Case "K�m�r Karas�"
        If index = 0 Then
                               RenkTemalar� = RGB(0, 0, 0)
        ElseIf index = 1 Then: RenkTemalar� = RGB(64, 64, 64)
        ElseIf index = 2 Then: RenkTemalar� = RGB(255, 255, 255)
        ElseIf index = 3 Then: RenkTemalar� = RGB(64, 64, 64)
        ElseIf index = 4 Then: RenkTemalar� = RGB(255, 255, 255)
        End If
    Case "Windows XP"
        If index = 0 Then
                               RenkTemalar� = -2147483633
        ElseIf index = 1 Then: RenkTemalar� = -2147483643
        ElseIf index = 2 Then: RenkTemalar� = -2147483640
        ElseIf index = 3 Then: RenkTemalar� = -2147483633
        ElseIf index = 4 Then: RenkTemalar� = -2147483640
        End If
    End Select
End Function
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+

