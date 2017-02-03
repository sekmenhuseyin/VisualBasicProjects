VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form Arama 
   BackColor       =   &H0096E06D&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7965
   ControlBox      =   0   'False
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   7965
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   4012
      Top             =   9240
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0096E06D&
      Caption         =   "Bulunan Kayýtlar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Left            =   270
      TabIndex        =   19
      Top             =   3222
      Width           =   7425
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3915
         IntegralHeight  =   0   'False
         ItemData        =   "frmSearch.frx":000C
         Left            =   210
         List            =   "frmSearch.frx":000E
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   525
         Width           =   6750
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(Adý,  Soyadý,  Müþteri Numarasý)"
         Height          =   330
         Left            =   225
         TabIndex        =   22
         Top             =   315
         Width           =   7125
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0096E06D&
      Caption         =   "Arama Kriterleri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   270
      TabIndex        =   18
      Top             =   192
      Width           =   7425
      Begin VB.ComboBox durum 
         Height          =   315
         ItemData        =   "frmSearch.frx":0010
         Left            =   5190
         List            =   "frmSearch.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1410
         Width           =   1785
      End
      Begin VB.TextBox tel 
         Height          =   285
         Left            =   5190
         TabIndex        =   11
         Top             =   1065
         Width           =   1785
      End
      Begin VB.CheckBox odurum 
         BackColor       =   &H0096E06D&
         Caption         =   "Ýþin Durumu"
         Height          =   255
         Left            =   3510
         TabIndex        =   12
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CheckBox otel 
         BackColor       =   &H0096E06D&
         Caption         =   "Telefon"
         Height          =   255
         Left            =   3510
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox osip 
         BackColor       =   &H0096E06D&
         Caption         =   "Sipariþ Tarihi"
         Height          =   255
         Left            =   255
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox opro 
         BackColor       =   &H0096E06D&
         Caption         =   "Prova Tarihi"
         Height          =   255
         Left            =   255
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox otes 
         BackColor       =   &H0096E06D&
         Caption         =   "Teslim Tarihi"
         Height          =   255
         Left            =   255
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox oadý 
         BackColor       =   &H0096E06D&
         Caption         =   "Müþterinin Adý"
         Height          =   255
         Left            =   3510
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox osoyadý 
         BackColor       =   &H0096E06D&
         Caption         =   "Müþterinin Soyadý"
         Height          =   255
         Left            =   3510
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox soyadý 
         Height          =   285
         Left            =   5190
         TabIndex        =   9
         Top             =   705
         Width           =   1785
      End
      Begin VB.TextBox adý 
         Height          =   285
         Left            =   5190
         TabIndex        =   7
         Top             =   345
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker sip 
         Height          =   285
         Left            =   1695
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CalendarBackColor=   16777215
         Format          =   7929857
         CurrentDate     =   38459
      End
      Begin MSComCtl2.DTPicker pro 
         Height          =   285
         Left            =   1695
         TabIndex        =   3
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   7929857
         CurrentDate     =   38459
      End
      Begin MSComCtl2.DTPicker tes 
         Height          =   285
         Left            =   1695
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   7929857
         CurrentDate     =   38459
      End
      Begin OsenXPCntrl.OsenXPButton Command2 
         Default         =   -1  'True
         Height          =   960
         Left            =   4560
         TabIndex        =   14
         Top             =   1845
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1693
         BTYPE           =   3
         TX              =   "Ara"
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
         MICON           =   "frmSearch.frx":002E
         PICN            =   "frmSearch.frx":004A
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Aranan:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   23
         Top             =   1680
         Width           =   675
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   3210
         X2              =   3210
         Y1              =   315
         Y2              =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   975
         Left            =   255
         TabIndex        =   21
         Top             =   1845
         Width           =   4365
      End
   End
   Begin OsenXPCntrl.OsenXPButton Command1 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   270
      TabIndex        =   17
      Top             =   8040
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
      MICON           =   "frmSearch.frx":0744
      PICN            =   "frmSearch.frx":0760
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton Command3 
      Height          =   735
      Left            =   5280
      TabIndex        =   16
      Top             =   8040
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
      MICON           =   "frmSearch.frx":0BB2
      PICN            =   "frmSearch.frx":0BCE
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox List2 
      Height          =   255
      ItemData        =   "frmSearch.frx":1020
      Left            =   3510
      List            =   "frmSearch.frx":1022
      TabIndex        =   24
      Top             =   4287
      Visible         =   0   'False
      Width           =   1275
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
      Left            =   2745
      TabIndex        =   20
      Top             =   8040
      Width           =   2475
   End
End
Attribute VB_Name = "Arama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command3_Click()    'Göster
    Write #7, Time$, Me.Name, "Command3_Click", "Start" 'logging
    On Local Error Resume Next
    If Trim(List1.Text) = "" Then Exit Sub
    Dim L_ID, tmp As String
    L_ID = ListeNO_Bul(List1.Text)
    If Label2.Caption = "(Adý,  Soyadý,  Müþteri Numarasý)" Then
        Open App.path + "\temp2.txt" For Output As #2: Write #2, L_ID: Close #2
        Dim i, j, k As Integer
        List1.Clear
        With Anasayfa.dtSipariþ.Recordset
            j = .RecordCount: If j = 0 Then Exit Sub
            .MoveFirst
            For i = 1 To j
                If Val(.Fields("Musteri_Kodu")) = Val(L_ID) Then
                    Anasayfa.dtSipariþTürü.Recordset.MoveFirst
                    For k = 1 To Anasayfa.dtSipariþTürü.Recordset.RecordCount
                        If Val(Anasayfa.dtSipariþTürü.Recordset.Fields("Siparis_Turleri")) = Val(.Fields("Siparis_Turu")) Then tmp = Anasayfa.dtSipariþTürü.Recordset.Fields("Siparis_Adi"): Exit For
                        Anasayfa.dtSipariþTürü.Recordset.MoveNext
                    Next k
                    List1.AddItem .Fields("Durum") & "    " & .Fields("Teslim_Tarihi") & "    " & tmp & "       " & .Fields("Ucret") & "      (" & .Fields("Siparis_Kodu") & ")"
                End If
                .MoveNext
            Next i
        End With
        Label2.Caption = "(Durum ,   Teslim Tarihi ,  Cinsi ,  Fiyatý ,  Kayýt Numarasý)"
        If List1.ListCount >= 1 Then List1.Selected(0) = True: List1.SetFocus: adet.Caption = List1.ListCount & " Kayýt Listede Görüntülendi." Else adet.Caption = "Görüntülenecek Kayýt Yok."
    Else
        Open App.path + "\temp2.txt" For Input As #2: Input #2, tmp: Close #2
        Open App.path + "\temp2.txt" For Output As #2: Write #2, tmp, L_ID: Close #2
        Me.Visible = False: Göster.Show: Göster.startingpoint.Caption = Me.Name
    End If
    Write #7, Time$, Me.Name, "Command3_Click", "End" 'logging
End Sub
Public Sub Command2_Click() 'ARA
    Dim sql As String
    Write #7, Time$, Me.Name, "Command2_Click", "Start" 'logging
    On Local Error Resume Next
    Me.Enabled = False: Me.MousePointer = 11
    'aramada bulunacak isimlerin tekrarlanmas için bulunanlar ilk önce bu dosyaya yazýlacak.
    Open App.path + "\temp.txt" For Output As #1:  Close #1
    'formda bir temizlik
    List1.Clear: List2.Clear: Label1.Caption = "": Label2.Caption = "(Adý,  Soyadý,  Müþteri Numarasý)": sql = ""
    'eðer bir seçilen bir kriterin içeriði boþsa o kriter iptal ediliyor.
    If adý.Text = "" Then oadý.Value = 0
    If soyadý.Text = "" Then osoyadý.Value = 0
    If tel.Text = "" Then otel.Value = 0
''''''''''''''''''''''''''''''''''''''''''''arama iþlemleri. seçilen bir arama kriteri varsa onu arayacak.''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''tbl_Musteriler
    If oadý.Value = 1 Then sql = sql + "[Musteri_Adi] like ""*" + Trim(CStr(adý.Text)) + "*"" or ": Label1.Caption = Label1.Caption & "Adý=" & CStr(adý.Text) & ",  "
    If osoyadý.Value = 1 Then sql = sql + "[Musteri_Soyadi] like ""*" + Trim(CStr(soyadý.Text)) + "*"" or ": Label1.Caption = Label1.Caption & "Soyadý=" & CStr(soyadý.Text) & ",  "
    If otel.Value = 1 Then sql = sql + "[Musteri_Telefon] like " + Trim(CStr(tel.Text)) + " or ": Label1.Caption = Label1.Caption & "Telefonu=" & CStr(tel.Text) & ",  "
    If oadý.Value = 1 Or osoyadý.Value = 1 Or otel.Value = 1 Then sql = VBA.Left(sql, Len(sql) - 4): Arama_ve_Ekleme 1, sql: sql = ""
    '''''''''''''''''''''''''tbl_Siparisler
    If otes.Value = 1 Then sql = sql + "[Teslim_Tarihi] like ""*" + CStr(tes.Value) + "*"" or ": Label1.Caption = Label1.Caption & "Teslim Tarihi=" & CStr(tes.Value) & ",  "
    If osip.Value = 1 Then sql = sql + "[Siparis_Tarihi] like ""*" + CStr(sip.Value) + "*"" or ": Label1.Caption = Label1.Caption & "Sipariþ Tarihi=" & CStr(sip.Value) & ",  "
    If opro.Value = 1 Then sql = sql + "[Prova_Tarihi] like ""*" + CStr(pro.Value) + "*"" or ": Label1.Caption = Label1.Caption & "Prova Tarihi=" & CStr(pro.Value) & ",  "
    If odurum.Value = 1 Then sql = sql + "[Durum] like ""*" + CStr(durum.Text) + "*"" or ": Label1.Caption = Label1.Caption & "Ýþin Durumu=" & CStr(durum.Text) & ",  "
    If otes.Value = 1 Or osip.Value = 1 Or opro.Value = 1 Then sql = VBA.Left(sql, Len(sql) - 4): Arama_ve_Ekleme 2, sql
    'arama bitti. þimdi arama raporu gibi yazýlar yazýlýyor. form kullanýlabilir bir hale geliyor.
    Anasayfa.dtMüþteri.RecordSource = "tbl_Musteriler": Anasayfa.dtMüþteri.Refresh: Anasayfa.dtSipariþ.RecordSource = "tbl_Siparisler": Anasayfa.dtSipariþ.Refresh
    If List1.ListCount >= 1 Then List1.Selected(0) = True: adet.Caption = List1.ListCount & " Kayýt Listede Görüntülendi." Else adet.Caption = "Görüntülenecek Kayýt Yok."
    If Label1.Caption <> "" Then Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 3)
    Anasayfa.Timer1.Enabled = True: Me.Enabled = True: Me.MousePointer = 1
    List1.SetFocus
    Write #7, Time$, Me.Name, "Command2_Click", "End" 'logging
End Sub
Sub Arama_ve_Ekleme(TabloAdý As String, Arama_Kriteri As String)
    Write #7, Time$, Me.Name, "Arama_ve_Ekleme", "Start" 'logging
    On Local Error Resume Next
    Dim Kimlik As String: Dim Önceden_Yazýlmýþ As Boolean: Dim i, j As Integer
    Select Case TabloAdý
        Case "1"
            Anasayfa.dtMüþteri.RecordSource = "tbl_Musteriler": Anasayfa.dtMüþteri.Refresh
            Arama_Kriteri = "select * from tbl_Musteriler where " + Arama_Kriteri: Anasayfa.dtMüþteri.RecordSource = Arama_Kriteri: Anasayfa.dtMüþteri.Refresh
            KayýtÝçinSayým 2: j = Anasayfa.dtMüþteri.Recordset.RecordCount: If j = 0 Then GoTo bitti
        Case "2"
            Anasayfa.dtSipariþ.RecordSource = "tbl_Siparisler": Anasayfa.dtSipariþ.Refresh
            Arama_Kriteri = "select * from tbl_Siparisler where " + Arama_Kriteri: Anasayfa.dtSipariþ.RecordSource = Arama_Kriteri: Anasayfa.dtSipariþ.Refresh
            KayýtÝçinSayým 7: j = Anasayfa.dtSipariþ.Recordset.RecordCount: If j = 0 Then GoTo bitti
    End Select
    For i = 1 To j
        Önceden_Yazýlmýþ = False
        If Anasayfa.dtMüþteri.Recordset.EOF = True Then GoTo bitti
        If Anasayfa.dtSipariþ.Recordset.EOF = True Then GoTo bitti
        Open App.path + "\temp.txt" For Input As #1
        Do While EOF(1) = False
            Input #1, Kimlik
            If TabloAdý = "1" Then
                If CStr(Kimlik) = CStr(Anasayfa.dtMüþteri.Recordset.Fields("Musteri_Kodu")) Then Önceden_Yazýlmýþ = True
            Else
                If CStr(Kimlik) = CStr(Anasayfa.dtSipariþ.Recordset.Fields("Musteri_Kodu")) Then Önceden_Yazýlmýþ = True
            End If
        Loop
        Close #1
        If Önceden_Yazýlmýþ = False Then
            Open App.path + "\temp.txt" For Append As #1
            If TabloAdý = "1" Then
                Write #1, Anasayfa.dtMüþteri.Recordset.Fields("Musteri_Kodu")
                List1.AddItem Anasayfa.dtMüþteri.Recordset.Fields("Musteri_Adi") & " " & Anasayfa.dtMüþteri.Recordset.Fields("Musteri_Soyadi") & String(5, " ") & "(" & Anasayfa.dtMüþteri.Recordset.Fields("Musteri_Kodu") & ")"
                List2.AddItem Anasayfa.dtMüþteri.Recordset.Fields("Musteri_Adi") & " " & Anasayfa.dtMüþteri.Recordset.Fields("Musteri_Soyadi") & String(5, " ") & "(" & Anasayfa.dtMüþteri.Recordset.Fields("Musteri_Kodu") & ")"
            Else
                Write #1, Anasayfa.dtSipariþ.Recordset.Fields("Musteri_Kodu")
                List1.AddItem GetMüþterifromID(Anasayfa.dtSipariþ.Recordset.Fields("Musteri_Kodu")) & String(5, " ") & "(" & Anasayfa.dtSipariþ.Recordset.Fields("Musteri_Kodu") & ")"
                List2.AddItem GetMüþterifromID(Anasayfa.dtSipariþ.Recordset.Fields("Musteri_Kodu")) & String(5, " ") & "(" & Anasayfa.dtSipariþ.Recordset.Fields("Musteri_Kodu") & ")"
            End If
            Close #1
        End If
        If TabloAdý = "1" Then Anasayfa.dtMüþteri.Recordset.MoveNext Else Anasayfa.dtSipariþ.Recordset.MoveNext
    Next i
bitti:
    Anasayfa.dtMüþteri.RecordSource = "tbl_Musteriler": Anasayfa.dtMüþteri.Refresh
    Anasayfa.dtSipariþ.RecordSource = "tbl_Siparisler": Anasayfa.dtSipariþ.Refresh
    Write #7, Time$, Me.Name, "Arama_ve_Ekleme", Arama_Kriteri 'logging
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub Form_Load()
    Write #7, Time$, Me.Name, "Form_Load", "Start" 'logging
    'boyamalar ve ilk ayarlamar
    Dim Control As Control: Me.BackColor = rnk_frm_arka
    For Each Control In Me
        If TypeOf Control Is OsenXPButton Then Control.BackColor = rnk_btn_arka: Control.BackOver = rnk_btn_arka: Control.ForeColor = rnk_btn_ön: Control.ForeOver = rnk_btn_ön
        If TypeOf Control Is Label Then Control.ForeColor = rnk_frm_ön
        If TypeOf Control Is Frame Then Control.BackColor = rnk_frm_arka: Control.ForeColor = rnk_frm_ön
        If TypeOf Control Is CheckBox Then Control.BackColor = rnk_frm_arka: Control.ForeColor = rnk_frm_ön
        If TypeOf Control Is ComboBox Then Control.BackColor = rnk_yazý_arka: Control.ForeColor = rnk_yazý_ön: Control.ListIndex = 0: Control.Enabled = False
        If TypeOf Control Is TextBox Then Control.BackColor = rnk_yazý_arka: Control.ForeColor = rnk_yazý_ön: Control.Enabled = False
        If TypeOf Control Is DTPicker Then Control.CalendarBackColor = rnk_yazý_arka: Control.CalendarForeColor = rnk_yazý_ön: Control.Enabled = False: Control.Value = Date
        If TypeOf Control Is ListBox Then Control.ForeColor = rnk_yazý_ön: Control.BackColor = rnk_yazý_arka
        If TypeOf Control Is Line Then Control.BorderColor = rnk_frm_ön
    Next Control
    otes.Value = 1: tes.Enabled = True: Open App.path + "\temp.txt" For Output As #1: Close #1
    Me.Show: frmMain.Caption = App.ProductName + "-Arama": Call frmMain.MDIForm_Resize
    Write #7, Time$, Me.Name, "Form_Load", "End" 'logging
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Write #7, Time$, Me.Name, "Form_Unload", "Successful" 'logging
    frmMain.Caption = App.ProductName: Anasayfa.Visible = True: Anasayfa.Command1.SetFocus
End Sub
Private Sub List1_DblClick()
  Command3_Click
End Sub
Private Sub List1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Command3_Click
End Sub
Private Sub Command1_Click()    'Geri
    If Label2.Caption = "(Adý,  Soyadý,  Müþteri Numarasý)" Then
        Unload Me
    Else
        Dim i As Integer: List1.Clear
        For i = 0 To List2.ListCount - 1
            List1.AddItem List2.List(i)
        Next i
        Label2.Caption = "(Adý,  Soyadý,  Müþteri Numarasý)"
        List1.SetFocus
    End If
End Sub
'+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
'formdaki textler checkler için ayrýlmýþ bir kod bölümü
'üzerine gelince tüm yazýlar seçiliyor.
Private Sub adý_GotFocus(): Call SelectAllText: End Sub
Private Sub tel_GotFocus(): Call SelectAllText: End Sub
Private Sub soyadý_GotFocus(): Call SelectAllText: End Sub
Private Sub tel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 32 Then Exit Sub
    If KeyAscii > 57 Or KeyAscii < 47 Then KeyAscii = 0
End Sub
'+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
'checkbox kodlarý

Private Sub odurum_Click()
    If odurum.Value = 1 Then durum.Enabled = True: durum.SetFocus Else durum.Enabled = False
End Sub
Private Sub otel_Click()
    If otel.Value = 1 Then tel.Enabled = True: tel.SetFocus Else tel.Enabled = False
End Sub
Private Sub oadý_Click()
    If oadý.Value = 1 Then adý.Enabled = True: adý.SetFocus Else adý.Enabled = False
End Sub
Private Sub osoyadý_Click()
    If osoyadý.Value = 1 Then soyadý.Enabled = True: soyadý.SetFocus Else soyadý.Enabled = False
End Sub
Private Sub opro_Click()
    If opro.Value = 1 Then pro.Enabled = True: pro.SetFocus Else pro.Enabled = False
End Sub
Private Sub osip_Click()
    If osip.Value = 1 Then sip.Enabled = True: sip.SetFocus Else sip.Enabled = False
End Sub
Private Sub otes_Click()
    If otes.Value = 1 Then tes.Enabled = True: tes.SetFocus Else tes.Enabled = False
End Sub
'+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
Private Sub Timer1_Timer()
    Command2_Click 'ilk aramamýz
    Timer1.Enabled = False
End Sub

