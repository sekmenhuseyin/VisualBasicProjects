Attribute VB_Name = "xmdOthers"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''en �stte
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const SWP_NOACTIVATE = &H10: Private Const SWP_SHOWWINDOW = &H40: Private Const HWND_NOTOPMOST = -2: Private Const HWND_TOPMOST = -1
Public Sub AlwaysOnTop(FormName As Form, SetOnTop As Boolean)
    Dim lFlag
    If SetOnTop Then lFlag = HWND_TOPMOST Else lFlag = HWND_NOTOPMOST
    SetWindowPos FormName.hwnd, lFlag, _
    FormName.Left / Screen.TwipsPerPixelX, FormName.Top / Screen.TwipsPerPixelY, _
    FormName.Width / Screen.TwipsPerPixelX, FormName.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub
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
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*

Public Function VarsaM��teriIDBul(M��teri_Ad� As String, M��teri_Soyad� As String) As String
    Dim i, j As Integer: Dim Cevap As String
    't�m m��teriler tablosu taranacak istenilen ad ve soyaddaki m��terinin kodunu bulmak i�in
    With Anasayfa.dtM��teri.Recordset
        Cevap = "0": j = .RecordCount
        If j = 0 Then VarsaM��teriIDBul = "0": Exit Function
        .MoveFirst
        For i = 1 To j
            If .Fields("Musteri_Adi") = M��teri_Ad� And .Fields("Musteri_Soyadi") = M��teri_Soyad� Then Cevap = .Fields("Musteri_Kodu"): Exit For
            .MoveNext
        Next i
        VarsaM��teriIDBul = Cevap
    End With
End Function
Public Function GetM��terifromID(AranacakM��etriKodu As String) As String
    Dim Kriter As String
    Kriter = "select * from tbl_Musteriler where [Musteri_Kodu]=" + CStr(AranacakM��etriKodu) + ""
    Anasayfa.dtM��teri.RecordSource = Kriter: Anasayfa.dtM��teri.Refresh: Anasayfa.dtM��teri.Recordset.MoveFirst
    GetM��terifromID = Anasayfa.dtM��teri.Recordset.Fields("Musteri_Adi") + " " + Anasayfa.dtM��teri.Recordset.Fields("Musteri_Soyadi")
    Anasayfa.dtM��teri.RecordSource = "tbl_Musteriler": Anasayfa.dtM��teri.Refresh
End Function
'g�ster formu i�in gerekli olan m��teri numaralar�n� bulacak
Public Function ListeNO_Bul(Aranacak�fade As String) As String
    Dim ID, tmp As String: Dim i As Integer
    ID = Left(Aranacak�fade, Len(Aranacak�fade) - 1)
    For i = Len(ID) To 1 Step -1
        If Mid(ID, i, 1) = "(" Then Exit For
    Next i
    ID = Right(ID, Len(ID) - i)
    ListeNO_Bul = ID
End Function
Public Sub Kay�t��inSay�m(DataNo As Byte)
    Select Case DataNo
        Case 2
            With Anasayfa.dtM��teri.Recordset
                Do While .EOF <> True: .MoveNext: Loop
                If .BOF <> True Then .MoveFirst
            End With
        Case 7
            With Anasayfa.dtSipari�.Recordset
                Do While .EOF <> True: .MoveNext: Loop
                If .BOF <> True Then .MoveFirst
            End With
    End Select
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
Public Sub SelectAllText()
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
End Sub

