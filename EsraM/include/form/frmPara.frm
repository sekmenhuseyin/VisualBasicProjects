VERSION 5.00
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form ParaGiri�i 
   BackColor       =   &H0096E06D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Para Giri�i"
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
      Caption         =   "Para Giri�i"
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
      Begin VB.TextBox al�nan 
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
         Caption         =   "Al�nan"
         Height          =   195
         Left            =   225
         TabIndex        =   14
         Top             =   420
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0096E06D&
      Caption         =   "M��teri Bilgileri"
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
      Begin VB.ComboBox ad� 
         Height          =   315
         Left            =   855
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   330
         Width           =   2445
      End
      Begin VB.ComboBox soyad� 
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
      Begin VB.Label soyad�2 
         Height          =   300
         Left            =   855
         TabIndex        =   17
         Top             =   817
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label ad�2 
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
         Caption         =   "Soyad�"
         Height          =   195
         Left            =   225
         TabIndex        =   12
         Top             =   870
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ad�"
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
Attribute VB_Name = "ParaGiri�i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GeriAl As Boolean
Private Sub Command1_Click()
    Write #7, Time$, Me.Name, "Command1_Click", "Start" 'logging
    On Local Error Resume Next
    Dim M��teriID As Long
    If Trim(ad�.Text) = "" Or Trim(soyad�.Text) = "" Or Val(al�nan.Text) = 0 Then MsgBox "L�tfen m��terinin ad�n�, soyad�n� ve al�nan miktar� do�ru giriniz !": Exit Sub
    If Trim(ad�2.Caption) <> "" Or Trim(soyad�2.Caption) <> "" Then
        M��teriID = Label7.Caption
    Else
        M��teriID = VarsaM��teriIDBul(ad�.Text, soyad�.Text)
        If M��teriID = 0 Then MsgBox "Bu ad ve soyada sahip bir m��teri bulunmamaktad�r !", vbCritical + vbOKOnly: Exit Sub
    End If
    If Val(kalan.Text) - Val(al�nan.Text) < 0 Then
        If MsgBox("M��terinin �demesi gerekenden daha fazla �cret al�nacak." & Chr(13) & "Yine de devam etmek istiyor musnuz?", vbYesNo + vbDefaultButton2, "�deme Fazlas�") = vbNo Then Exit Sub
    End If
    With Anasayfa.dt�deme.Recordset
        .AddNew
        .Fields("Musteri_Kodu") = M��teriID
        .Fields("Odenen_Fiyat") = al�nan.Text
        .Fields("Tarih") = tarih.Text
        .Update
    End With
    kalan.Text = Val(kalan.Text) - Val(al�nan.Text): al�nan.Text = ""
    Write #7, Time$, Me.Name, "Command1_Click", "Successful" 'logging
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub ad�_GotFocus(): Call SelectAllText: End Sub
Private Sub fiyat_GotFocus(): Call SelectAllText: End Sub
Private Sub kalan_GotFocus(): Call SelectAllText: End Sub
Private Sub soyad�_GotFocus(): Call SelectAllText: End Sub
Private Sub al�nan_GotFocus(): Call SelectAllText: End Sub
Private Sub tarih_GotFocus(): Call SelectAllText: End Sub
Private Sub ad�_Change()
    Dim i As Long: Dim nSel As Long
    If GeriAl = True Or ad�.Text = "" Then GeriAl = False: Exit Sub
    For i = 0 To ad�.ListCount - 1
        If InStr(1, ad�.List(i), ad�.Text, vbTextCompare) = 1 Then
            nSel = ad�.SelStart: ad�.Text = ad�.List(i): ad�.SelStart = nSel: ad�.SelLength = Len(ad�.Text) - nSel
            Exit For
        End If
    Next
End Sub
Private Sub ad�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If ad�.Text <> "" Then GeriAl = True
    End If
End Sub
Private Sub ad�_LostFocus()
    If Trim(ad�2.Caption) <> "" Or Trim(soyad�2.Caption) <> "" Then Exit Sub
    Dim i As Integer
    soyad�.Clear
    With Anasayfa.dtM��teri.Recordset
        .MoveFirst
        For i = 1 To .RecordCount
            If ad�.Text = .Fields("Musteri_Adi") Then soyad�.AddItem .Fields("Musteri_Soyadi")
            .MoveNext
        Next i
    End With
End Sub
Private Sub soyad�_Change()
    Dim i As Long: Dim nSel As Long
    If GeriAl = True Or soyad�.Text = "" Then GeriAl = False: Exit Sub
    For i = 0 To soyad�.ListCount - 1
        If InStr(1, soyad�.List(i), soyad�.Text, vbTextCompare) = 1 Then
            nSel = soyad�.SelStart: soyad�.Text = soyad�.List(i): soyad�.SelStart = nSel: soyad�.SelLength = Len(soyad�.Text) - nSel
            Exit For
        End If
    Next
End Sub
Private Sub soyad�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If soyad�.Text <> "" Then GeriAl = True
    End If
End Sub
Private Sub soyad�_LostFocus()
    Dim M��teriID As Long: Dim temp As String: Dim i As Integer
    fiyat.Text = "0": kalan.Text = "0": temp = "0"
    If Trim(ad�2.Caption) <> "" Or Trim(soyad�2.Caption) <> "" Then
        M��teriID = Anasayfa.dtM��teri.Recordset.Fields("Musteri_Kodu")
        temp = CStr(Anasayfa.dtSipari�.Recordset.Fields("Siparis_Kodu"))
    Else
        M��teriID = VarsaM��teriIDBul(ad�.Text, soyad�.Text)
    End If
    With Anasayfa.dtSipari�.Recordset 'sipari�ler
        .MoveFirst
        For i = 1 To .RecordCount
            If Val(.Fields("Musteri_Kodu")) = Val(M��teriID) Then fiyat.Text = Val(fiyat.Text) + Val(.Fields("Ucret"))
            .MoveNext
        Next i
    End With: DoEvents
    With Anasayfa.dt�deme.Recordset '�demeler
        If .RecordCount = 0 Then GoTo son
        .MoveFirst
        For i = 1 To .RecordCount
            If Val(.Fields("Musteri_Kodu")) = Val(M��teriID) Then kalan.Text = Val(kalan.Text) + Val(.Fields("Odenen_Fiyat"))
            .MoveNext
        Next i
    End With
son:
    If Trim(ad�2.Caption) <> "" Or Trim(soyad�2.Caption) <> "" Then
        With Anasayfa.dtSipari�.Recordset 'sipari�ler
            .MoveFirst: While CStr(.Fields("Siparis_Kodu")) <> temp: .MoveNext: Wend
        End With
    End If
    kalan.Text = Val(fiyat.Text) - Val(kalan.Text)
End Sub
Private Sub al�nan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 13 Then Exit Sub
    If KeyAscii > 57 Or KeyAscii < 47 Then KeyAscii = 0
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub Timer1_Timer()
    Write #7, Time$, Me.Name, "Timer1_Timer", "Start" 'logging
    Dim i, j As Integer: Dim M��teriName As String
    With Anasayfa.dtM��teri.Recordset
        If Trim(ad�2.Caption) <> "" And Trim(soyad�2.Caption) <> "" Then
            ad�.AddItem .Fields("Musteri_Adi"): ad�.ListIndex = 0: ad�.Locked = True: soyad�.SetFocus
            soyad�.AddItem .Fields("Musteri_Soyadi"): soyad�.ListIndex = 0: soyad�.Locked = True: al�nan.SetFocus
        Else
            .MoveFirst
            For i = 1 To .RecordCount
                M��teriName = .Fields("Musteri_Adi")
                For j = 0 To ad�.ListCount
                    If UCase(ad�.List(j)) = UCase(M��teriName) Then GoTo bas
                Next j
                ad�.AddItem M��teriName
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
        If TypeOf Control Is OsenXPButton Then Control.BackColor = rnk_btn_arka: Control.BackOver = rnk_btn_arka: Control.ForeColor = rnk_btn_�n: Control.ForeOver = rnk_btn_�n
        If TypeOf Control Is Label Then Control.ForeColor = rnk_frm_�n
        If TypeOf Control Is Frame Then Control.BackColor = rnk_frm_arka: Control.ForeColor = rnk_frm_�n
        If TypeOf Control Is ComboBox Then Control.BackColor = rnk_yaz�_arka: Control.ForeColor = rnk_yaz�_�n
        If TypeOf Control Is TextBox Then Control.BackColor = rnk_yaz�_arka: Control.ForeColor = rnk_yaz�_�n
        If TypeOf Control Is Line Then Control.BorderColor = rnk_frm_�n
    Next Control
    Write #7, Time$, Me.Name, "Form_Load", "Successful" 'logging
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If startingpoint.Caption = "G�ster" Then
        G�ster.Enabled = True
    ElseIf startingpoint.Caption = "Malii�lemler" Then
        Malii�lemler.Enabled = True
    End If
    Write #7, Time$, Me.Name, "Form_Unload", "Successful" 'logging
End Sub


