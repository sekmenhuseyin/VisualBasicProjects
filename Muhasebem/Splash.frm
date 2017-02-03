VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00CDB75F&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4515
   ClientLeft      =   2295
   ClientTop       =   2475
   ClientWidth     =   7605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   90
      Top             =   105
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3945
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   7080
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "Splash.frx":030A
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright 2004© Sekmenler Tech."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4500
         TabIndex        =   3
         Top             =   3060
         Width           =   2475
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   " Warning: This software is licensed for the end-user using only ."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   4
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2520
         TabIndex        =   6
         Top             =   1140
         Width           =   2430
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sekmenler Tech."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2355
         TabIndex        =   5
         Top             =   705
         Width           =   2850
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   7
      Height          =   3945
      Left            =   150
      Top             =   165
      Width           =   7080
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   15
      Height          =   3825
      Left            =   360
      Top             =   420
      Width           =   6975
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Dosya1, Dosya2, Dosya3, Dosya4, Dosya5, Dosya6, Dosya7
Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
'*****************olmasý gereken dosyalar kontrol ediliyor*********************
'-----------------------1
Dosya1 = Dir(App.Path + "\data\cmd1.nfo")
If Dosya1 = "" Then
'markalar
Open App.Path + "\data\cmd1.nfo" For Append As #1
Close #1
End If
'-----------------------2
Dosya2 = Dir(App.Path + "\data\cmd3.nfo")
If Dosya2 = "" Then
'modeller
Open App.Path + "\data\cmd3.nfo" For Append As #1
Close #1
End If
'-----------------------3
Dosya3 = Dir(App.Path + "\data\cmd5.nfo")
If Dosya3 = "" Then
'türler
Open App.Path + "\data\cmd5.nfo" For Append As #1
Close #1
End If
'-----------------------4
Dosya4 = Dir(App.Path + "\data\cmd7.nfo")
If Dosya4 = "" Then
'garantiler
Open App.Path + "\data\cmd7.nfo" For Append As #1
Close #1
End If
'-----------------------5
Dosya5 = Dir(App.Path + "\stuff\stok.mlz")
If Dosya5 = "" Then
'miktarý öðrenmek,malzeme listesi
Open App.Path + "\stuff\stok.mlz" For Append As #1
Close #1
End If
'-----------------------6
Dosya6 = Dir(App.Path + "\stuff\market.mlz")
If Dosya6 = "" Then
'malzeme satýþý için bilgi depolamak
Open App.Path + "\stuff\market.mlz" For Append As #1
Close #1
End If
'-----------------------7
Dosya7 = Dir(App.Path + "\stuff\hst.mlz")
If Dosya7 = "" Then
'stok kontrolü için gerekli -hem malzeme giriþi hem de çýkýþý yazýlacak-
Open App.Path + "\stuff\hst.mlz" For Append As #1
Close #1
End If
frmSplash.Left = Screen.Width / 2 - 3817: frmSplash.Top = Screen.Height / 2 - 2272
Frame1.Left = frmSplash.Width / 2 - 3660: Frame1.Top = frmSplash.Height / 2 - 2109
Shape1.Left = Frame1.Left: Shape1.Top = Frame1.Top
Shape2.Left = Frame1.Left + 210: Shape2.Top = Frame1.Top + 255
End Sub
Private Sub Timer1_Timer()
Form1.Enabled = True
Unload Me
End Sub
