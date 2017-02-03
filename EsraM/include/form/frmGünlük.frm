VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form Gsipariþ 
   BackColor       =   &H0096E06D&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7965
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmGünlük.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   7965
   Begin VB.ListBox List3 
      Height          =   1620
      ItemData        =   "frmGünlük.frx":000C
      Left            =   8760
      List            =   "frmGünlük.frx":000E
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2160
      Width           =   3495
   End
   Begin VB.ListBox List2 
      Height          =   1620
      ItemData        =   "frmGünlük.frx":0010
      Left            =   8760
      List            =   "frmGünlük.frx":0012
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   9975
      Top             =   6930
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0096E06D&
      Caption         =   "Günlük Takip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6300
      Left            =   191
      TabIndex        =   3
      Top             =   1530
      Width           =   7575
      Begin VB.ListBox List1 
         Height          =   5115
         IntegralHeight  =   0   'False
         ItemData        =   "frmGünlük.frx":0014
         Left            =   135
         List            =   "frmGünlük.frx":0016
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   990
         Width           =   7155
      End
      Begin VB.ListBox List_Gizli 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5115
         IntegralHeight  =   0   'False
         ItemData        =   "frmGünlük.frx":0018
         Left            =   135
         List            =   "frmGünlük.frx":001A
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   990
         Visible         =   0   'False
         Width           =   7155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Durum       Ýþ tipi       Kumas -  Cinsi            Ad Soyad     Kayýt No)"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   750
         Width           =   4590
      End
      Begin VB.Label durum 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   7575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0096E06D&
      Caption         =   "Ýþ Tarihleri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   191
      TabIndex        =   1
      Top             =   180
      Width           =   7575
      Begin OsenXPCntrl.OsenXPButton Command3 
         Default         =   -1  'True
         Height          =   855
         Left            =   5040
         TabIndex        =   12
         Top             =   255
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1508
         BTYPE           =   3
         TX              =   "Takvime Bak"
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
         MICON           =   "frmGünlük.frx":001C
         PICN            =   "frmGünlük.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker tarih 
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   495
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   661
         _Version        =   393216
         Format          =   8126465
         CurrentDate     =   38459
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ýþleri Görmek Ýstediðiniz Tarih :"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   585
         Width           =   2145
      End
   End
   Begin VB.Timer gecikme 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5933
      Top             =   533
   End
   Begin VB.ListBox List4 
      Height          =   1620
      ItemData        =   "frmGünlük.frx":0912
      Left            =   8790
      List            =   "frmGünlük.frx":0914
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4080
      Width           =   3495
   End
   Begin OsenXPCntrl.OsenXPButton Command1 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   195
      TabIndex        =   13
      Top             =   7995
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
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
      MICON           =   "frmGünlük.frx":0916
      PICN            =   "frmGünlük.frx":0932
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton Command2 
      Height          =   735
      Left            =   5355
      TabIndex        =   14
      Top             =   7995
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Göster"
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
      MICON           =   "frmGünlük.frx":0D84
      PICN            =   "frmGünlük.frx":0DA0
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label takvimon 
      BackStyle       =   0  'Transparent
      Caption         =   "no"
      Height          =   435
      Left            =   3255
      TabIndex        =   16
      Top             =   3780
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Hangisi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   8790
      TabIndex        =   9
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label adet 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2595
      TabIndex        =   5
      Top             =   7995
      Width           =   2745
   End
End
Attribute VB_Name = "Gsipariþ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Gör_Click()
    Write #7, Time$, Me.Name, "Gör_Click", "Start" 'logging
    On Local Error Resume Next
    Dim i, j As Integer: Dim tar As Date: Dim temp, temp2, MüþteriAdSoyad As String
    Hangisi.Caption = "0": tar = tarih.Value: durum.Caption = "Görüntülenen Tarih : " & tar
    List_Gizli.Visible = True: List1.Clear: List2.Clear: List3.Clear
    If tar = Date Then durum.Caption = durum.Caption & " (Bugün)" Else If tar + 1 = Date Then durum.Caption = durum.Caption & " (Dün)"
    Label2.Caption = "(Durum       Ýþ tipi       Kumas -  Cinsi            Ad Soyad     Kayýt No)"
    With Anasayfa.dtSipariþ.Recordset
        If .RecordCount = 0 Then GoTo bitti
        .MoveFirst
        For i = 1 To .RecordCount
            If tar = .Fields("Siparis_Tarihi") Or tar = .Fields("Prova_Tarihi") Or tar = .Fields("Teslim_Tarihi") Then
                Anasayfa.dtSipariþTürü.Recordset.MoveFirst
                For j = 1 To Anasayfa.dtSipariþTürü.Recordset.RecordCount
                    If Val(Anasayfa.dtSipariþTürü.Recordset.Fields("Siparis_Turleri")) = Val(.Fields("Siparis_Turu")) Then temp = Anasayfa.dtSipariþTürü.Recordset.Fields("Siparis_Adi"): Exit For
                    Anasayfa.dtSipariþTürü.Recordset.MoveNext
                Next j
                Anasayfa.dtKumaþTürü.Recordset.MoveFirst
                For j = 1 To Anasayfa.dtKumaþTürü.Recordset.RecordCount
                    If Val(Anasayfa.dtKumaþTürü.Recordset.Fields("Kumas_Turu")) = Val(.Fields("Kumas_Turu")) Then temp2 = Anasayfa.dtKumaþTürü.Recordset.Fields("Kumas_Adi"): Exit For
                    Anasayfa.dtKumaþTürü.Recordset.MoveNext
                Next j
                MüþteriAdSoyad = GetMüþterifromID(.Fields("Musteri_Kodu"))
            End If
            If tar = .Fields("Siparis_Tarihi") Then
                List2.AddItem "(" & .Fields("Durum") & ")   Sipariþ       " & temp2 & "  " & temp & Space(10) & MüþteriAdSoyad & "    (" & .Fields("Siparis_Kodu") & ")"
                List3.AddItem .Fields("Musteri_Kodu")
                List1.AddItem List2.List(List2.ListCount - 1)
            End If
            If tar = .Fields("Prova_Tarihi") Then
                List2.AddItem "(" & .Fields("Durum") & ")   Prova        " & temp2 & "  " & temp & Space(10) & MüþteriAdSoyad & "    (" & .Fields("Siparis_Kodu") & ")"
                List3.AddItem .Fields("Musteri_Kodu")
                List1.AddItem List2.List(List2.ListCount - 1)
            End If
            If tar = .Fields("Teslim_Tarihi") Then
                List2.AddItem "(" & .Fields("Durum") & ")   Teslim       " & temp2 & "  " & temp & Space(10) & MüþteriAdSoyad & "    (" & .Fields("Siparis_Kodu") & ")"
                List3.AddItem .Fields("Musteri_Kodu")
                List1.AddItem List2.List(List2.ListCount - 1)
            End If
            .MoveNext: DoEvents
        Next i
    End With
bitti:
    If List1.ListCount >= 1 Then List1.Selected(0) = True: adet.Caption = List1.ListCount & " Kayýt Listede Görüntülendi." Else adet.Caption = "Görüntülenecek Kayýt Yok."
    List_Gizli.Visible = False: If takvimon.Caption <> "yes" Then List1.SetFocus
    Write #7, Time$, Me.Name, "Gör_Click", tarih.Value 'logging
End Sub
Private Sub Command2_Click()
    Write #7, Time$, Me.Name, "Command2_Click", "Start" 'logging
    If Trim(List1.Text) = "" Then Exit Sub
    Dim asd As Date: Dim i As Integer
    If Hangisi.Caption = "0" Then
        For i = 0 To List1.ListCount
            If List1.Text = List2.List(i) Then Exit For
        Next i
        Open App.path + "\temp2.txt" For Output As #2
            Write #2, List3.List(i)
            Write #2, ListeNO_Bul(List1.Text)
        Close #2
        If takvimon.Caption = "yes" Then Unload Takvim
        Me.Visible = False: Göster.Show: Göster.startingpoint.Caption = Me.Name
    Else
        asd = Left(List1.Text, 10)
        tarih.Value = asd
        Gör_Click
    End If
    Write #7, Time$, Me.Name, "Command2_Click", "End" 'logging
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub tarih_Change()
    Gör_Click
End Sub
Private Sub Timer1_Timer()
    Gör_Click
    Timer1.Enabled = False
End Sub
Private Sub Command1_Click()
    If takvimon.Caption = "yes" Then Unload Takvim
    Unload Me
End Sub
Private Sub List1_DblClick()
    Command2_Click
End Sub
Private Sub Command3_Click()
    Takvim.Show: takvimon.Caption = "yes"
End Sub
Private Sub Form_Load()
    Write #7, Time$, Me.Name, "Form_Load", "Start" 'logging
    Dim Control As Control: Me.BackColor = rnk_frm_arka
    For Each Control In Me
        If TypeOf Control Is OsenXPButton Then Control.BackColor = rnk_btn_arka: Control.BackOver = rnk_btn_arka: Control.ForeColor = rnk_btn_ön: Control.ForeOver = rnk_btn_ön
        If TypeOf Control Is Label Then Control.ForeColor = rnk_frm_ön
        If TypeOf Control Is Frame Then Control.BackColor = rnk_frm_arka: Control.ForeColor = rnk_frm_ön
        If TypeOf Control Is DTPicker Then Control.CalendarBackColor = rnk_yazý_arka: Control.CalendarForeColor = rnk_yazý_ön: Control.Value = Date
        If TypeOf Control Is ListBox Then Control.ForeColor = rnk_yazý_ön: Control.BackColor = rnk_yazý_arka
    Next Control
    Me.Show: frmMain.Caption = App.ProductName + "-Günlük Sipariþ": Call frmMain.MDIForm_Resize: Timer1.Enabled = True
    If List1.ListCount >= 1 Then adet.Caption = List1.ListCount & " Kayýt Listede Görüntülendi." Else adet.Caption = "Görüntülenecek Kayýt Yok."
    Write #7, Time$, Me.Name, "Form_Load", "Successful" 'logging
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmMain.Caption = App.ProductName: Anasayfa.Visible = True: Anasayfa.Command1.SetFocus
    Write #7, Time$, Me.Name, "Form_Unload", "Successful" 'logging
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+



