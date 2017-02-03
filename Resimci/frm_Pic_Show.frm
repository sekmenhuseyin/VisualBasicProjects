VERSION 5.00
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form Photo_Show 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Resim Gösterici"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   Icon            =   "frm_Pic_Show.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame_BG 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10455
      Begin VB.VScrollBar VScroll 
         Enabled         =   0   'False
         Height          =   6600
         LargeChange     =   500
         Left            =   10200
         SmallChange     =   100
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.HScrollBar HScroll 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   500
         Left            =   0
         SmallChange     =   100
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   6600
         Visible         =   0   'False
         Width           =   10215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   10185
         TabIndex        =   10
         Top             =   6615
         Width           =   1065
      End
      Begin VB.Label Label_Err 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resim Görüntülenemedi"
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Image Picture_Main 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   6585
         Left            =   0
         Top             =   0
         Width           =   10185
      End
   End
   Begin VB.Frame Frame_Details 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3930
      TabIndex        =   4
      Top             =   6840
      Width           =   2685
      Begin OsenXPCntrl.OsenXPButton Command1 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   105
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   "Sýðdýr"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_Pic_Show.frx":2372
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton Command2 
         Height          =   375
         Left            =   740
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   105
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   "Geniþlet"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_Pic_Show.frx":238E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton ZoomPlus 
         Height          =   375
         Left            =   1740
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_Pic_Show.frx":23AA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton ZoomMinus 
         Height          =   375
         Left            =   2120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_Pic_Show.frx":23C6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         DrawMode        =   14  'Copy Pen
         X1              =   1575
         X2              =   1575
         Y1              =   105
         Y2              =   525
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   540
      Left            =   4620
      TabIndex        =   0
      Top             =   9450
      Width           =   1170
   End
End
Attribute VB_Name = "Photo_Show"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ratio As Single
Private Sub Command1_Click()
    VScroll.Enabled = False: HScroll.Enabled = False
    VScroll.Visible = False: HScroll.Visible = False
    Picture_Main.Stretch = True
    If ratio < 1 Then   'geniþlik daha uzun
            'geniþliðe göre ayarlarsak yükseklik iyi gelecekmi
            If ((Me.Width - 375) * ratio) > Me.Height - 1350 Then 'hatalý. yüksekliðe göre ayarlanacak
                Picture_Main.Height = Me.Height - 1350
                Picture_Main.Width = Picture_Main.Height / ratio
            Else
                Picture_Main.Width = Me.Width - 375
                Picture_Main.Height = Picture_Main.Width * ratio
            End If
    Else 'ratio>1 yani  'yükseklik daha uzun
            'yüksekliðe göre ayarlarsak geniþlik iyi gelecekmi
            If ((Me.Height - 1350) / ratio) > Me.Width - 375 Then 'hatalý. geniþliðe göre ayarlanacak
                Picture_Main.Width = Me.Width - 375
                Picture_Main.Height = Picture_Main.Width * ratio
            Else
                Picture_Main.Height = Me.Height - 1350
                Picture_Main.Width = Picture_Main.Height / ratio
            End If
    End If
    If Picture_Main.Width < Me.Width - 375 Then Picture_Main.Left = Me.Width / 2 - Picture_Main.Width / 2 - 140 Else Picture_Main.Left = 0
    If Picture_Main.Height < Me.Height - 1350 Then Picture_Main.Top = Me.Height / 2 - Picture_Main.Height / 2 - 675 Else Picture_Main.Top = 0
    Command1.Enabled = False: Command2.Enabled = True: Command3.SetFocus
End Sub
Private Sub Command2_Click()
    Picture_Main.Stretch = False
    Komand2nin_YerAyarlamalarý
    Command1.Enabled = True: Command2.Enabled = False: Command3.SetFocus
End Sub
Sub Komand2nin_YerAyarlamalarý()
    On Error Resume Next
    'resmin geniþliði ile ilgili ayarlamalar
    If Picture_Main.Width > Me.Width - 375 Then
        If Picture_Main.Width - Frame_BG.Width > 0 Then HScroll.Max = Picture_Main.Width - Frame_BG.Width Else HScroll.Max = 0
        Picture_Main.Left = -HScroll.Value
        If HScroll.Max = 0 Then HScroll.Enabled = False: HScroll.Visible = False Else HScroll.Enabled = True: HScroll.Visible = True
    Else
        HScroll.Enabled = False: HScroll.Visible = False
        Picture_Main.Left = Me.Width / 2 - Picture_Main.Width / 2 - 140
    End If
    'resmin yüksekliði ile ilgili ayarlamalar
    If Picture_Main.Height > Me.Height - 1350 Then
        If Picture_Main.Height - Frame_BG.Height > 0 Then VScroll.Max = Picture_Main.Height - Frame_BG.Height Else VScroll.Max = 0
        Picture_Main.Top = -VScroll.Value
        If VScroll.Max = 0 Then VScroll.Enabled = False: VScroll.Visible = False Else VScroll.Enabled = True: VScroll.Visible = True
    Else
        VScroll.Enabled = False: VScroll.Visible = False
        Picture_Main.Top = Me.Height / 2 - Picture_Main.Height / 2 - 675
    End If
End Sub
Private Sub ZoomMinus_Click()
    Picture_Main.Stretch = True
    Picture_Main.Height = Picture_Main.Height \ 1.5
    Picture_Main.Width = Picture_Main.Height / ratio
    Komand2nin_YerAyarlamalarý
    If ((Picture_Main.Height / 1.5) < 500) Then ZoomMinus.Enabled = False
    If ((Picture_Main.Width / 1.5) < 500) Then ZoomMinus.Enabled = False
    Command1.Enabled = True: Command2.Enabled = True
    ZoomPlus.Enabled = True: Command3.SetFocus
End Sub
Private Sub ZoomPlus_Click()
    Picture_Main.Stretch = True
    Picture_Main.Height = Picture_Main.Height * 1.5
    Picture_Main.Width = Picture_Main.Height / ratio
    Komand2nin_YerAyarlamalarý
    If ((Picture_Main.Height * 1.5) - Frame_BG.Height) > 32500 Then ZoomPlus.Enabled = False
    If ((Picture_Main.Width * 1.5) - Frame_BG.Width) > 32500 Then ZoomPlus.Enabled = False
    Command1.Enabled = True: Command2.Enabled = True
    ZoomMinus.Enabled = True: Command3.SetFocus
End Sub
Sub Non_Changable_Telif_Check()
    If App.CompanyName = "Sekmenler Tech." Then
        If App.LegalCopyright = "© " + App.CompanyName Then Exit Sub
    End If
    MsgBox "Uygulamanýn telif haklarý deðiþtirilmiþ." + Chr(13) + Chr(10) + "Lütfen uygulamayý tekrar kurun."
    End
End Sub
Private Sub Form_Load()
    'telif haklarý kontrolü yapýlýyor!
    Non_Changable_Telif_Check
    'hata olduðunda yani resim yüklenmediðinde...
    On Error GoTo to_Error_Service
    If Trim(Command) <> "" Then
        Picture_Main.Picture = LoadPicture(Command)
        ratio = Picture_Main.Height / Picture_Main.Width
    Else
        GoTo to_Error_Service  'resim gösterilemiyor diye yazdýrýlacak orada...
    End If
    Exit Sub
to_Error_Service:
    If Trim(Command) <> "" Then Label_Err.Caption = "Resim Görüntülenemedi" Else Label_Err.Caption = "Resim Belirtilmedi"
    Error_Service
    Label_Err.Visible = True    'hata yazýsýnýn olduðu label!
End Sub
Private Sub Error_Service()
    Label_Err.Top = Frame_BG.Height / 2
    Label_Err.Left = Frame_BG.Width / 2 - Label_Err.Width / 2 - 52
    Command1.Enabled = False: Command2.Enabled = False
    ZoomPlus.Enabled = False: ZoomMinus.Enabled = False
    Picture_Main.Visible = False
End Sub
Private Sub Form_Resize()
    'engellemeler (belli bir boyuttan daha küçük olamaz...)
    If Me.Width < 5370 And Me.Height < 6150 Then
        Me.Width = 5370: Me.Height = 6150 ': Exit Sub
    ElseIf Me.Height < 6150 Then
        Me.Height = 6150 ': Exit Sub
    ElseIf Me.Width < 5370 Then
        Me.Width = 5370 ': Exit Sub
    End If
    Frame_BG.Move 0, 0, ScaleWidth, ScaleHeight - Frame_Details.Height
    VScroll.Move ScaleWidth - VScroll.Width, 0, VScroll.Width, ScaleHeight - Frame_Details.Height - HScroll.Height
    HScroll.Move 0, ScaleHeight - HScroll.Height - Frame_Details.Height, ScaleWidth - VScroll.Width
    If ScaleWidth < Frame_Details.Width Then
        Frame_Details.Move 0, HScroll.Top + HScroll.Height
    Else
        Frame_Details.Move (ScaleWidth - Frame_Details.Width) / 2, HScroll.Top + HScroll.Height
    End If
    Label1.Move HScroll.Width, VScroll.Height
    'Diðer ayarlar
    If Label_Err.Visible = True Then
        Error_Service
    Else
        If Command1.Enabled = False And Command2.Enabled = True Then
            Command1_Click
        ElseIf Command1.Enabled = True And Command2.Enabled = False Then
            Command2_Click
        ElseIf Command1.Enabled = True And Command2.Enabled = True Then
            Komand2nin_YerAyarlamalarý
        End If
    End If
End Sub
Private Sub HScroll_Change()
    Picture_Main.Left = -HScroll.Value
End Sub
Private Sub VScroll_Change()
    Picture_Main.Top = -VScroll.Value
End Sub

