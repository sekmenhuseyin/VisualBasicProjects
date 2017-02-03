Attribute VB_Name = "xmdOthers"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''en üstte
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
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*

Public Function VarsaMüþteriIDBul(Müþteri_Adý As String, Müþteri_Soyadý As String) As String
    Dim i, j As Integer: Dim Cevap As String
    'tüm müþteriler tablosu taranacak istenilen ad ve soyaddaki müþterinin kodunu bulmak için
    With Anasayfa.dtMüþteri.Recordset
        Cevap = "0": j = .RecordCount
        If j = 0 Then VarsaMüþteriIDBul = "0": Exit Function
        .MoveFirst
        For i = 1 To j
            If .Fields("Musteri_Adi") = Müþteri_Adý And .Fields("Musteri_Soyadi") = Müþteri_Soyadý Then Cevap = .Fields("Musteri_Kodu"): Exit For
            .MoveNext
        Next i
        VarsaMüþteriIDBul = Cevap
    End With
End Function
Public Function GetMüþterifromID(AranacakMüþetriKodu As String) As String
    Dim Kriter As String
    Kriter = "select * from tbl_Musteriler where [Musteri_Kodu]=" + CStr(AranacakMüþetriKodu) + ""
    Anasayfa.dtMüþteri.RecordSource = Kriter: Anasayfa.dtMüþteri.Refresh: Anasayfa.dtMüþteri.Recordset.MoveFirst
    GetMüþterifromID = Anasayfa.dtMüþteri.Recordset.Fields("Musteri_Adi") + " " + Anasayfa.dtMüþteri.Recordset.Fields("Musteri_Soyadi")
    Anasayfa.dtMüþteri.RecordSource = "tbl_Musteriler": Anasayfa.dtMüþteri.Refresh
End Function
'göster formu için gerekli olan müþteri numaralarýný bulacak
Public Function ListeNO_Bul(AranacakÝfade As String) As String
    Dim ID, tmp As String: Dim i As Integer
    ID = Left(AranacakÝfade, Len(AranacakÝfade) - 1)
    For i = Len(ID) To 1 Step -1
        If Mid(ID, i, 1) = "(" Then Exit For
    Next i
    ID = Right(ID, Len(ID) - i)
    ListeNO_Bul = ID
End Function
Public Sub KayýtÝçinSayým(DataNo As Byte)
    Select Case DataNo
        Case 2
            With Anasayfa.dtMüþteri.Recordset
                Do While .EOF <> True: .MoveNext: Loop
                If .BOF <> True Then .MoveFirst
            End With
        Case 7
            With Anasayfa.dtSipariþ.Recordset
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

