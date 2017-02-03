VERSION 5.00
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form ParaGiriþi 
   BackColor       =   &H0096E06D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Para Giriþi"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   Icon            =   "frmPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2730
      Top             =   4095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0096E06D&
      Caption         =   "Para Giriþi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   97
      TabIndex        =   13
      Top             =   1553
      Width           =   5910
      Begin VB.TextBox tarih 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4200
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox alýnan 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   855
         TabIndex        =   2
         Top             =   360
         Width           =   2445
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tarih"
         Height          =   195
         Left            =   3690
         TabIndex        =   15
         Top             =   420
         Width           =   360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   3510
         X2              =   3510
         Y1              =   240
         Y2              =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alýnan"
         Height          =   195
         Left            =   225
         TabIndex        =   14
         Top             =   420
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0096E06D&
      Caption         =   "Müþteri Bilgileri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   97
      TabIndex        =   6
      Top             =   98
      Width           =   5910
      Begin VB.ComboBox adý 
         Height          =   315
         Left            =   855
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   330
         Width           =   2445
      End
      Begin VB.ComboBox soyadý 
         Height          =   315
         Left            =   855
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   810
         Width           =   2445
      End
      Begin VB.TextBox fiyat 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4230
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   330
         Width           =   1455
      End
      Begin VB.TextBox kalan 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4230
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   810
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   105
         TabIndex        =   18
         Top             =   210
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label soyadý2 
         Height          =   300
         Left            =   855
         TabIndex        =   17
         Top             =   817
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label adý2 
         Height          =   300
         Left            =   855
         TabIndex        =   16
         Top             =   337
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Soyadý"
         Height          =   195
         Left            =   225
         TabIndex        =   12
         Top             =   870
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adý"
         Height          =   195
         Left            =   225
         TabIndex        =   11
         Top             =   390
         Width           =   225
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tutar"
         Height          =   195
         Left            =   3690
         TabIndex        =   10
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kalan"
         Height          =   195
         Left            =   3690
         TabIndex        =   9
         Top             =   870
         Width           =   405
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   3510
         X2              =   3510
         Y1              =   240
         Y2              =   1200
      End
   End
   Begin OsenXPCntrl.OsenXPButton Command2 
      Cancel          =   -1  'True
      Height          =   840
      Left            =   97
      TabIndex        =   4
      Top             =   2543
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1482
      BTYPE           =   3
      TX              =   "Geri"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmPara.frx":0442
      PICN            =   "frmPara.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton Command1 
      Default         =   -1  'True
      Height          =   840
      Left            =   3592
      TabIndex        =   3
      Top             =   2543
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1482
      BTYPE           =   3
      TX              =   "Kaydet"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmPara.frx":08B0
      PICN            =   "frmPara.frx":08CC
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label startingpoint 
      BackStyle       =   0  'Transparent
      Height          =   540
      Left            =   2415
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "ParaGiriþi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GeriAl As Boolean
Private Sub Command1_Click()
    Write #7, Time$, Me.Name, "Command1_Click", "Start" 'logging
    On Local Error Resume Next
    Dim MüþteriID As Long
    If Trim(adý.Text) = "" Or Trim(soyadý.Text) = "" Or Val(alýnan.Text) = 0 Then MsgBox "Lütfen müþterinin adýný, soyadýný ve alýnan miktarý doðru giriniz !": Exit Sub
    If Trim(adý2.Caption) <> "" Or Trim(soyadý2.Caption) <> "" Then
        MüþteriID = Label7.Caption
    Else
        MüþteriID = VarsaMüþteriIDBul(adý.Text, soyadý.Text)
        If MüþteriID = 0 Then MsgBox "Bu ad ve soyada sahip bir müþteri bulunmamaktadýr !", vbCritical + vbOKOnly: Exit Sub
    End If
    If Val(kalan.Text) - Val(alýnan.Text) < 0 Then
        If MsgBox("Müþterinin ödemesi gerekenden daha fazla ücret alýnacak." & Chr(13) & "Yine de devam etmek istiyor musnuz?", vbYesNo + vbDefaultButton2, "Ödeme Fazlasý") = vbNo Then Exit Sub
    End If
    With Anasayfa.dtÖdeme.Recordset
        .AddNew
        .Fields("Musteri_Kodu") = MüþteriID
        .Fields("Odenen_Fiyat") = alýnan.Text
        .Fields("Tarih") = tarih.Text
        .Update
    End With
    kalan.Text = Val(kalan.Text) - Val(alýnan.Text): alýnan.Text = ""
    Write #7, Time$, Me.Name, "Command1_Click", "Successful" 'logging
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub adý_GotFocus(): Call SelectAllText: End Sub
Private Sub fiyat_GotFocus(): Call SelectAllText: End Sub
Private Sub kalan_GotFocus(): Call SelectAllText: End Sub
Private Sub soyadý_GotFocus(): Call SelectAllText: End Sub
Private Sub alýnan_GotFocus(): Call SelectAllText: End Sub
Private Sub tarih_GotFocus(): Call SelectAllText: End Sub
Private Sub adý_Change()
    Dim i As Long: Dim nSel As Long
    If GeriAl = True Or adý.Text = "" Then GeriAl = False: Exit Sub
    For i = 0 To adý.ListCount - 1
        If InStr(1, adý.List(i), adý.Text, vbTextCompare) = 1 Then
            nSel = adý.SelStart: adý.Text = adý.List(i): adý.SelStart = nSel: adý.SelLength = Len(adý.Text) - nSel
            Exit For
        End If
    Next
End Sub
Private Sub adý_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If adý.Text <> "" Then GeriAl = True
    End If
End Sub
Private Sub adý_LostFocus()
    If Trim(adý2.Caption) <> "" Or Trim(soyadý2.Caption) <> "" Then Exit Sub
    Dim i As Integer
    soyadý.Clear
    With Anasayfa.dtMüþteri.Recordset
        .MoveFirst
        For i = 1 To .RecordCount
            If adý.Text = .Fields("Musteri_Adi") Then soyadý.AddItem .Fields("Musteri_Soyadi")
            .MoveNext
        Next i
    End With
End Sub
Private Sub soyadý_Change()
    Dim i As Long: Dim nSel As Long
    If GeriAl = True Or soyadý.Text = "" Then GeriAl = False: Exit Sub
    For i = 0 To soyadý.ListCount - 1
        If InStr(1, soyadý.List(i), soyadý.Text, vbTextCompare) = 1 Then
            nSel = soyadý.SelStart: soyadý.Text = soyadý.List(i): soyadý.SelStart = nSel: soyadý.SelLength = Len(soyadý.Text) - nSel
            Exit For
        End If
    Next
End Sub
Private Sub soyadý_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If soyadý.Text <> "" Then GeriAl = True
    End If
End Sub
Private Sub soyadý_LostFocus()
    Dim MüþteriID As Long: Dim temp As String: Dim i As Integer
    fiyat.Text = "0": kalan.Text = "0": temp = "0"
    If Trim(adý2.Caption) <> "" Or Trim(soyadý2.Caption) <> "" Then
        MüþteriID = Anasayfa.dtMüþteri.Recordset.Fields("Musteri_Kodu")
        temp = CStr(Anasayfa.dtSipariþ.Recordset.Fields("Siparis_Kodu"))
    Else
        MüþteriID = VarsaMüþteriIDBul(adý.Text, soyadý.Text)
    End If
    With Anasayfa.dtSipariþ.Recordset 'sipariþler
        .MoveFirst
        For i = 1 To .RecordCount
            If Val(.Fields("Musteri_Kodu")) = Val(MüþteriID) Then fiyat.Text = Val(fiyat.Text) + Val(.Fields("Ucret"))
            .MoveNext
        Next i
    End With: DoEvents
    With Anasayfa.dtÖdeme.Recordset 'ödemeler
        If .RecordCount = 0 Then GoTo son
        .MoveFirst
        For i = 1 To .RecordCount
            If Val(.Fields("Musteri_Kodu")) = Val(MüþteriID) Then kalan.Text = Val(kalan.Text) + Val(.Fields("Odenen_Fiyat"))
            .MoveNext
        Next i
    End With
son:
    If Trim(adý2.Caption) <> "" Or Trim(soyadý2.Caption) <> "" Then
        With Anasayfa.dtSipariþ.Recordset 'sipariþler
            .MoveFirst: While CStr(.Fields("Siparis_Kodu")) <> temp: .MoveNext: Wend
        End With
    End If
    kalan.Text = Val(fiyat.Text) - Val(kalan.Text)
End Sub
Private Sub alýnan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 13 Then Exit Sub
    If KeyAscii > 57 Or KeyAscii < 47 Then KeyAscii = 0
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub Timer1_Timer()
    Write #7, Time$, Me.Name, "Timer1_Timer", "Start" 'logging
    Dim i, j As Integer: Dim MüþteriName As String
    With Anasayfa.dtMüþteri.Recordset
        If Trim(adý2.Caption) <> "" And Trim(soyadý2.Caption) <> "" Then
            adý.AddItem .Fields("Musteri_Adi"): adý.ListIndex = 0: adý.Locked = True: soyadý.SetFocus
            soyadý.AddItem .Fields("Musteri_Soyadi"): soyadý.ListIndex = 0: soyadý.Locked = True: alýnan.SetFocus
        Else
            .MoveFirst
            For i = 1 To .RecordCount
                MüþteriName = .Fields("Musteri_Adi")
                For j = 0 To adý.ListCount
                    If UCase(adý.List(j)) = UCase(MüþteriName) Then GoTo bas
                Next j
                adý.AddItem MüþteriName
bas:
                .MoveNext
            Next i
        End If
    End With
    Write #7, Time$, Me.Name, "Timer1", "Successful" 'logging
    Timer1.Enabled = False
End Sub
Private Sub Form_Load()
    Write #7, Time$, Me.Name, "Form_Load", "Start" 'logging
    Dim i As Integer: Dim Control As Control
    Me.BackColor = rnk_frm_arka: tarih.Text = Date
    For Each Control In Me
        If TypeOf Control Is OsenXPButton Then Control.BackColor = rnk_btn_arka: Control.BackOver = rnk_btn_arka: Control.ForeColor = rnk_btn_ön: Control.ForeOver = rnk_btn_ön
        If TypeOf Control Is Label Then Control.ForeColor = rnk_frm_ön
        If TypeOf Control Is Frame Then Control.BackColor = rnk_frm_arka: Control.ForeColor = rnk_frm_ön
        If TypeOf Control Is ComboBox Then Control.BackColor = rnk_yazý_arka: Control.ForeColor = rnk_yazý_ön
        If TypeOf Control Is TextBox Then Control.BackColor = rnk_yazý_arka: Control.ForeColor = rnk_yazý_ön
        If TypeOf Control Is Line Then Control.BorderColor = rnk_frm_ön
    Next Control
    Write #7, Time$, Me.Name, "Form_Load", "Successful" 'logging
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If startingpoint.Caption = "Göster" Then
        Göster.Enabled = True
    ElseIf startingpoint.Caption = "Maliiþlemler" Then
        Maliiþlemler.Enabled = True
    End If
    Write #7, Time$, Me.Name, "Form_Unload", "Successful" 'logging
End Sub


