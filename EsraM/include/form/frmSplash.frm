VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4230
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7365
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lütfen Bekleyiniz"
         Height          =   195
         Left            =   3015
         TabIndex        =   9
         Top             =   3660
         Width           =   1200
      End
      Begin VB.Image Image1 
         Height          =   1110
         Left            =   3120
         Picture         =   "frmSplash.frx":0442
         Top             =   840
         Width           =   3555
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":0AC1
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright: Sekmenler Tech."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4560
         TabIndex        =   4
         Top             =   3060
         Width           =   1980
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company: Sekmen Programlama"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4560
         TabIndex        =   3
         Top             =   3270
         Width           =   2310
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Temalar araþtýrýlýyor..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   3660
         Width           =   1545
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sürüm 1.1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5040
         TabIndex        =   5
         Top             =   2400
         Width           =   1260
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lisans Alýnmýþtýr."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dikim Evi"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   21.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3690
         TabIndex        =   6
         Top             =   1800
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackColor       =   &H0096E06D&
         Height          =   3875
         Left            =   45
         TabIndex        =   7
         Top             =   120
         Width           =   6970
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   315
      TabIndex        =   8
      Top             =   420
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5985
      Top             =   2835
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Goster As Integer
Private Sub Form_Load()
    Write #7, Time$, Me.Name, "Form_Load", "Start" 'logging
    On Local Error Resume Next
    Dim theme_name As String: Dim i, j As Byte: Dim Control As Control
    Me.BackColor = rnk_Btn_Arka: Goster = 0: Timer1.Enabled = True
    For Each Control In Me 'formu renklendireceðiz.
        If TypeOf Control Is OsenXPButton Then Control.BackColor = rnk_Btn_Arka: Control.BackOver = rnk_Btn_Arka: Control.ForeColor = rnk_Btn_Ön: Control.ForeOver = rnk_Btn_Ön
        If TypeOf Control Is Label Then Control.ForeColor = rnk_Frm_Ön: Control.BackColor = rnk_Frm_Arka
        If TypeOf Control Is Frame Then Control.BackColor = rnk_Btn_Arka
    Next Control
    'þimdi de var olan temalar belleðe aktarýlýyor.
    'seçili olan tema bulunduðunda ise o temanýn resimleri forma yükleniyor!
    Dir1.path = App.path + "\Style\": ReDim Themes(Dir1.ListCount)
    For i = 1 To Dir1.ListCount
        Dir1.ListIndex = Dir1.ListIndex + 1
        For j = 0 To Len(Dir1.List(Dir1.ListIndex))
            If Mid(Dir1.List(Dir1.ListIndex), Len(Dir1.List(Dir1.ListIndex)) - j, 1) = "\" Then theme_name = Right(Dir1.List(Dir1.ListIndex), j): Exit For
        Next j
        Themes(Dir1.ListIndex).TemaAd = theme_name: Themes(Dir1.ListIndex).TemaDizin = Dir1.List(Dir1.ListIndex)
        If theme_name = Tema_Adý Then Tema_Yeri = Dir1.List(Dir1.ListIndex)
    Next i
    If Dir(Tema_Yeri & "\logo.gif") <> "" Then Image1.Picture = LoadPicture(Tema_Yeri & "\logo.gif")
    lblVersion.Caption = "Sürüm " & App.Major & "." & App.Minor & "." & App.Revision
    lblCompany.Caption = App.CompanyName: lblCopyright = App.LegalCopyright
    AlwaysOnTop Me, True
    'logging
    Write #7, Time$, Me.Name, "Form_Load", "Successful"
End Sub
Private Sub Timer1_Timer()
  Select Case Goster
    Case 0
        Me.Show
    Case 1
        lblWarning.Caption = "Baþlatýlýyor..."
    Case 2
        Load frmMain
    Case 3
        Load Anasayfa: Anasayfa.Show: Call AddMenu
    Case 4
        Timer1.Enabled = False: Me.Hide
        Write #7, Time$, Me.Name, "Form_Unload", "Successful" 'logging
        Unload Me
  End Select
  Goster = Goster + 1
End Sub
Private Sub AddMenu()
    Write #7, Time$, Me.Name, "AddMenu", "Start:" & Dir1.ListCount 'logging
    Dim Index, i As Integer
    Anasayfa.altTema(0).Caption = Themes(0).TemaAd
    If Dir1.ListCount <> 0 Then
        For i = 1 To Dir1.ListCount - 1
            Index = Anasayfa.altTema.Count: Load Anasayfa.altTema(Index) 'yeni menu oluþturuluyor...
            Anasayfa.altTema(Index).Caption = Themes(i).TemaAd: Anasayfa.altTema(Index).Visible = True
            If Anasayfa.altTema(Index).Caption = Tema_Adý Then Anasayfa.altTema(Index).Checked = True
        Next i
    End If
    Write #7, Time$, Me.Name, "AddMenu", "End"
End Sub

