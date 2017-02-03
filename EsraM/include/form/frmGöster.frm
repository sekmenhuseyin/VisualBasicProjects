VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form Göster 
   BackColor       =   &H0096E06D&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
   ClientLeft      =   1065
   ClientTop       =   1845
   ClientWidth     =   7965
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmGöster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   7965
   Begin VB.Frame Frame4 
      BackColor       =   &H0096E06D&
      Caption         =   "Ayrýntýlar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Left            =   180
      TabIndex        =   49
      Top             =   6795
      Width           =   7575
      Begin OsenXPCntrl.OsenXPButton Command7 
         Height          =   450
         Left            =   225
         TabIndex        =   21
         Top             =   300
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   794
         BTYPE           =   3
         TX              =   "Resmi Göster"
         ENAB            =   0   'False
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
         MICON           =   "frmGöster.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton Command8 
         Height          =   450
         Left            =   6150
         TabIndex        =   22
         Top             =   300
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   794
         BTYPE           =   3
         TX              =   "Ýþ Bitti"
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
         MICON           =   "frmGöster.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   3780
         X2              =   3780
         Y1              =   210
         Y2              =   850
      End
      Begin VB.Label durumne 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3885
         TabIndex        =   50
         Top             =   405
         Width           =   2235
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1230
      Left            =   180
      TabIndex        =   39
      Top             =   270
      Width           =   7575
      Begin VB.TextBox adý 
         Height          =   285
         Left            =   1395
         TabIndex        =   3
         Top             =   360
         Width           =   2200
      End
      Begin VB.TextBox soyadý 
         Height          =   285
         Left            =   5115
         TabIndex        =   4
         Top             =   360
         Width           =   2200
      End
      Begin VB.TextBox fiyat 
         Height          =   285
         Left            =   5115
         TabIndex        =   6
         Top             =   720
         Width           =   2200
      End
      Begin VB.TextBox tel 
         Height          =   285
         Left            =   1395
         TabIndex        =   5
         Top             =   720
         Width           =   2200
      End
      Begin VB.Label adý2 
         Height          =   300
         Left            =   1260
         TabIndex        =   52
         Top             =   315
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label soyadý2 
         Height          =   300
         Left            =   4935
         TabIndex        =   51
         Top             =   315
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   3765
         X2              =   3765
         Y1              =   360
         Y2              =   1000
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fiyatý (YTL)"
         Height          =   195
         Left            =   3885
         TabIndex        =   43
         Top             =   765
         Width           =   795
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefonu"
         Height          =   195
         Left            =   225
         TabIndex        =   42
         Top             =   765
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Soyadý"
         Height          =   195
         Left            =   3885
         TabIndex        =   41
         Top             =   405
         Width           =   480
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adý"
         Height          =   195
         Left            =   225
         TabIndex        =   40
         Top             =   405
         Width           =   225
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0096E06D&
      Caption         =   "Mali Bilgiler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Left            =   180
      TabIndex        =   44
      Top             =   1605
      Width           =   7575
      Begin OsenXPCntrl.OsenXPButton Command9 
         Height          =   450
         Left            =   210
         TabIndex        =   7
         Top             =   300
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   794
         BTYPE           =   3
         TX              =   "Ödemeleri Göster"
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
         MICON           =   "frmGöster.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton Command10 
         Height          =   450
         Left            =   3990
         TabIndex        =   8
         Top             =   300
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   794
         BTYPE           =   3
         TX              =   "Para Giriþi"
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
         MICON           =   "frmGöster.frx":0060
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   3780
         X2              =   3780
         Y1              =   210
         Y2              =   850
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0096E06D&
      Caption         =   " Ýþ Bilgileri "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   180
      TabIndex        =   26
      Top             =   2670
      Width           =   7575
      Begin VB.ComboBox kumas 
         Height          =   315
         ItemData        =   "frmGöster.frx":007C
         Left            =   5115
         List            =   "frmGöster.frx":007E
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   2200
      End
      Begin VB.ComboBox cins 
         Height          =   315
         ItemData        =   "frmGöster.frx":0080
         Left            =   5115
         List            =   "frmGöster.frx":0082
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   2200
      End
      Begin VB.TextBox acik 
         Height          =   1890
         Left            =   3880
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   1845
         Width           =   3435
      End
      Begin VB.TextBox bel 
         Height          =   285
         Left            =   1395
         TabIndex        =   12
         Top             =   1650
         Width           =   2200
      End
      Begin VB.TextBox basen 
         Height          =   285
         Left            =   1395
         TabIndex        =   13
         Top             =   2010
         Width           =   2200
      End
      Begin VB.TextBox gogus 
         Height          =   285
         Left            =   1395
         TabIndex        =   14
         Top             =   2370
         Width           =   2200
      End
      Begin VB.TextBox omuz 
         Height          =   285
         Left            =   1395
         TabIndex        =   15
         Top             =   2730
         Width           =   2200
      End
      Begin VB.TextBox kol 
         Height          =   285
         Left            =   1395
         TabIndex        =   16
         Top             =   3090
         Width           =   2200
      End
      Begin VB.TextBox boy 
         Height          =   285
         Left            =   1395
         TabIndex        =   17
         Top             =   3450
         Width           =   2200
      End
      Begin MSComCtl2.DTPicker sip 
         Height          =   315
         Left            =   1395
         TabIndex        =   9
         Top             =   360
         Width           =   2200
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Format          =   7995393
         CurrentDate     =   38459
      End
      Begin MSComCtl2.DTPicker pro 
         Height          =   315
         Left            =   1395
         TabIndex        =   10
         Top             =   720
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Format          =   7995393
         CurrentDate     =   38459
      End
      Begin MSComCtl2.DTPicker tes 
         Height          =   315
         Left            =   1395
         TabIndex        =   11
         Top             =   1080
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Format          =   7995393
         CurrentDate     =   38459
      End
      Begin VB.Label cins2 
         Height          =   300
         Left            =   4935
         TabIndex        =   48
         Top             =   735
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label kumas2 
         Height          =   300
         Left            =   4935
         TabIndex        =   47
         Top             =   315
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Açýklama   :"
         Height          =   195
         Left            =   3885
         TabIndex        =   38
         Top             =   1590
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sipariþ Tarihi"
         Height          =   195
         Left            =   225
         TabIndex        =   37
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prova Tarihi"
         Height          =   195
         Left            =   225
         TabIndex        =   36
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teslim Tarihi"
         Height          =   195
         Left            =   225
         TabIndex        =   35
         Top             =   1140
         Width           =   885
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kumaþ Türü"
         Height          =   195
         Left            =   3885
         TabIndex        =   34
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Elbise Cinsi"
         Height          =   195
         Left            =   3885
         TabIndex        =   33
         Top             =   780
         Width           =   795
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   3765
         X2              =   3765
         Y1              =   375
         Y2              =   3780
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   150
         X2              =   7245
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bel"
         Height          =   195
         Left            =   225
         TabIndex        =   32
         Top             =   1695
         Width           =   225
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basen"
         Height          =   195
         Left            =   225
         TabIndex        =   31
         Top             =   2055
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Göðüs"
         Height          =   195
         Left            =   225
         TabIndex        =   30
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Omuz"
         Height          =   195
         Left            =   225
         TabIndex        =   29
         Top             =   2775
         Width           =   405
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kol"
         Height          =   195
         Left            =   210
         TabIndex        =   28
         Top             =   3135
         Width           =   225
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Boy"
         Height          =   195
         Left            =   225
         TabIndex        =   27
         Top             =   3495
         Width           =   270
      End
   End
   Begin VB.Frame FR_BTN1 
      BackColor       =   &H0096E06D&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   855
      Left            =   180
      TabIndex        =   45
      Top             =   7845
      Width           =   7605
      Begin OsenXPCntrl.OsenXPButton Command1 
         Cancel          =   -1  'True
         Height          =   855
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   1508
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
         MICON           =   "frmGöster.frx":0084
         PICN            =   "frmGöster.frx":00A0
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton Command5 
         Height          =   855
         Left            =   2535
         TabIndex        =   1
         Top             =   0
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   1508
         BTYPE           =   3
         TX              =   "Kaydý Sil"
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
         MICON           =   "frmGöster.frx":04F2
         PICN            =   "frmGöster.frx":050E
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
         Default         =   -1  'True
         Height          =   855
         Left            =   5130
         TabIndex        =   2
         Top             =   0
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   1508
         BTYPE           =   3
         TX              =   "Düzenle"
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
         MICON           =   "frmGöster.frx":0960
         PICN            =   "frmGöster.frx":097C
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame FR_BTN2 
      BackColor       =   &H0096E06D&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   180
      TabIndex        =   46
      Top             =   7845
      Visible         =   0   'False
      Width           =   7605
      Begin OsenXPCntrl.OsenXPButton Command6 
         Height          =   855
         Left            =   3
         TabIndex        =   25
         Top             =   0
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   1508
         BTYPE           =   3
         TX              =   "Ýptal"
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
         MICON           =   "frmGöster.frx":1256
         PICN            =   "frmGöster.frx":1272
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton Command4 
         Height          =   855
         Left            =   2535
         TabIndex        =   24
         Top             =   0
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   1508
         BTYPE           =   3
         TX              =   "Sýfýrla"
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
         MICON           =   "frmGöster.frx":16C4
         PICN            =   "frmGöster.frx":16E0
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
         Height          =   855
         Left            =   5130
         TabIndex        =   23
         Top             =   0
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   1508
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
         MICON           =   "frmGöster.frx":1B32
         PICN            =   "frmGöster.frx":1B4E
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Label startingpoint 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   540
      Left            =   3315
      TabIndex        =   53
      Top             =   3645
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "Göster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GeriAl, Deðiþtimi As Boolean
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub Command1_Click()    'geri
    Unload Me
End Sub
Private Sub Command5_Click()    'kaydý sil
    On Local Error Resume Next
    If MsgBox("Bu kaydý silmek istediðinizden emin misinizi?", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub
    Write #7, Time$, Me.Name, "Command5_Click", "Start" 'logging
    Anasayfa.dtSipariþ.Recordset.Delete: Deðiþtimi = True
    Write #7, Time$, Me.Name, "Command5_Click", "sil" 'logging
    Unload Me
End Sub
Private Sub Command3_Click()    'düzenle
    Write #7, Time$, Me.Name, "Command3_Click", "Start" 'logging
    Dim Control As Control
    For Each Control In Me
        If TypeOf Control Is TextBox Then Control.Locked = False
        If TypeOf Control Is ComboBox Then Control.Locked = False
        If TypeOf Control Is DTPicker Then Control.Enabled = True
    Next Control
    FR_BTN1.Visible = False: FR_BTN2.Visible = True: Command6.Cancel = True: Command2.Default = True: sip.SetFocus
    Write #7, Time$, Me.Name, "Command3_Click", "düzenle" 'logging
End Sub
Private Sub Command6_Click()    'iptal
    Write #7, Time$, Me.Name, "Command6_Click", "Start" 'logging
    Dim Control As Control
    Call Command4_Click
    For Each Control In Me
        If TypeOf Control Is TextBox Then Control.Locked = True
        If TypeOf Control Is ComboBox Then Control.Locked = True
        If TypeOf Control Is DTPicker Then Control.Enabled = False
    Next Control
    FR_BTN1.Visible = True: FR_BTN2.Visible = False: Command1.Cancel = True: Command3.Default = True: Command1.SetFocus
    Write #7, Time$, Me.Name, "Command6_Click", "iptal" 'logging
End Sub
Private Sub Command4_Click()    'sýfýrla
    Write #7, Time$, Me.Name, "Command4_Click", "Start" 'logging
    On Local Error Resume Next
    With Anasayfa.dtMüþteri.Recordset
        adý.Text = .Fields("Musteri_Adi")
        soyadý.Text = .Fields("Musteri_Soyadi")
        tel.Text = .Fields("Musteri_Telefon")
    End With
    With Anasayfa.dtSipariþ.Recordset
        fiyat.Text = .Fields("Ucret")
        acik.Text = .Fields("Aciklama")
        sip.Value = .Fields("Siparis_Tarihi")
        pro.Value = .Fields("Prova_Tarihi")
        tes.Value = .Fields("Teslim_Tarihi")
        boy.Text = .Fields("Boy")
        basen.Text = .Fields("Basen")
        bel.Text = .Fields("Bel")
        kol.Text = .Fields("Kol")
        gogus.Text = .Fields("Göðüs")
        omuz.Text = .Fields("Omuz")
    End With
    kumas.Text = kumas2.Caption
    cins.Text = cins2.Caption
    sip.SetFocus
    Write #7, Time$, Me.Name, "Command4_Click", "sýfýrla" 'logging
End Sub
Private Sub Command2_Click()    'kaydet
    Write #7, Time$, Me.Name, "Command2_Click", "Start" 'logging
    On Local Error Resume Next
    If fiyat = "" Or Val(fiyat) = 0 Then MsgBox "Lütfen malýn fiyatýný yazýnýz.": Exit Sub
    Dim Control As Control: Dim BoþYerVarMý, tmp_Sayý As Boolean: Dim i As Integer
    Dim MüþteriID, SipTürID, KumaþTürID As String
    BoþYerVarMý = False
    For Each Control In Me
        If TypeOf Control Is TextBox Then
            Control.Text = Trim(Control.Text)
            If Control.Text = "" Then BoþYerVarMý = True
        End If
    Next Control
    If BoþYerVarMý = True Then
        If MsgBox("Eksik Bilgi Ýçeriyor Devam Edilsin mi?", 36) = vbNo Then Exit Sub
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''KAYIT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ilk önce müþteri tablosu
    With Anasayfa.dtMüþteri.Recordset
        If adý.Text <> adý2.Caption Or soyadý.Text <> soyadý2.Caption Then
            If MsgBox("Müþterinin adýný ve soyadýný deðiþtirmek istediðinizden emin misiniz?", vbYesNo + vbDefaultButton2, "Müþteri Kayýt") = vbYes Then
                MüþteriID = VarsaMüþteriIDBul(adý.Text, soyadý.Text)
                If MüþteriID = 0 Then
                    .AddNew
                    .Fields("Musteri_Adi") = adý.Text
                    .Fields("Musteri_Soyadi") = soyadý.Text
                    .Update
                    .MoveLast
                    MüþteriID = Val(.Fields("Musteri_Kodu"))
                End If
            Else
                Open App.path + "\temp2.txt" For Input As #2: Input #2, MüþteriID: Close #2
            End If
        End If
        .Edit
        .Fields("Musteri_Telefon") = tel.Text
        .Update
    End With
    'sipariþ Türleri
    tmp_Sayý = False
    With Anasayfa.dtSipariþTürü.Recordset
        If cins.Text <> cins2.Caption Then
            If .RecordCount <> 0 Then
                .MoveFirst
                For i = 1 To .RecordCount  'burada bu sipariþ türünün daha önce kaydolup olmadýðýný araþtýrýyoruz.
                    If .Fields("Siparis_Adi") = cins.Text Then tmp_Sayý = True: Exit For 'kaydolmuþsa tmp_sayý=1 oluyor!
                    .MoveNext
                Next i
            End If
            If tmp_Sayý = False Then
                .AddNew
                .Fields("Siparis_Adi") = cins.Text
                .Update
                .MoveLast
                SipTürID = Val(.Fields("Siparis_Turleri"))
            Else
                SipTürID = Val(.Fields("Siparis_Turleri"))
            End If
        Else
            SipTürID = Val(.Fields("Siparis_Turleri"))
        End If
    End With
    'kumaþ Türleri
    tmp_Sayý = False
    With Anasayfa.dtKumaþTürü.Recordset
        If kumas.Text <> kumas2.Caption Then
            If .RecordCount <> 0 Then
                .MoveFirst
                For i = 1 To .RecordCount  'burada bu kumaþ türünün daha önce kaydolup olmadýðýný araþtýrýyoruz.
                    If .Fields("Kumas_Adi") = kumas.Text Then tmp_Sayý = True: Exit For 'kaydolmuþsa tmp_sayý=1 oluyor!
                    .MoveNext
                Next i
            End If
            If tmp_Sayý = False Then
                .AddNew
                .Fields("Kumas_Adi") = kumas.Text
                .Update
                .MoveLast
                KumaþTürID = Val(.Fields("Kumas_Turu"))
            Else
                KumaþTürID = Val(.Fields("Kumas_Turu"))
            End If
        Else
            KumaþTürID = Val(.Fields("Kumas_Turu"))
        End If
    End With
    'sýra geldi asýl iþe, yani sipariþ datasýna...
    With Anasayfa.dtSipariþ.Recordset
        .Edit
        .Fields("Musteri_Kodu") = MüþteriID
        .Fields("Siparis_Turu") = SipTürID
        .Fields("Kumas_Turu") = KumaþTürID
        .Fields("Siparis_Tarihi") = sip.Value
        .Fields("Prova_Tarihi") = pro.Value
        .Fields("Teslim_Tarihi") = tes.Value
        .Fields("Aciklama") = acik.Text
        .Fields("Ucret") = fiyat.Text
        .Fields("Durum") = "Bitmedi"
        .Fields("Boy") = boy.Text
        .Fields("Basen") = basen.Text
        .Fields("Bel") = bel.Text
        .Fields("Göðüs") = gogus.Text
        .Fields("Omuz") = omuz.Text
        .Fields("Kol") = kol.Text
        .Update
    End With
    If adý.Text <> adý2.Caption Or soyadý.Text <> soyadý2.Caption Then
        Dim tmpMüþteri, tmpSipariþ As String
        Open App.path + "\temp2.txt" For Input As #2: Input #2, tmpMüþteri: Input #2, tmpSipariþ: Close #2
        With Anasayfa.dtÖdeme.Recordset 'ödemeler
            If .RecordCount <> 0 Then
                .MoveFirst
                For i = 0 To .RecordCount
                    If .Fields("Musteri_Kodu") = tmpMüþteri Then
                        .Edit
                        .Fields("Musteri_Kodu") = MüþteriID
                        .Update
                    End If
                Next i
            End If
        End With
        With Anasayfa.dtResim.Recordset 'resimler
            If .RecordCount <> 0 Then
                .MoveFirst
                For i = 0 To .RecordCount
                    If .Fields("Siparis_Kodu") = tmpSipariþ Then
                        .Edit
                        .Fields("Siparis_Kodu") = Anasayfa.dtSipariþ.Recordset.Fields("Siparis_Kodu")
                        .Update
                    End If
                Next i
            End If
        End With
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''KAYIT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each Control In Me
        If TypeOf Control Is TextBox Then Control.Locked = True
        If TypeOf Control Is ComboBox Then Control.Locked = True
        If TypeOf Control Is DTPicker Then Control.Enabled = False
    Next Control
    FR_BTN1.Visible = True: FR_BTN2.Visible = False: Command1.Cancel = True: Command3.Default = True: Deðiþtimi = True
    durumne.Caption = "Durum :    Ýþ Bitmedi": Command8.Visible = True: Command1.SetFocus
    Write #7, Time$, Me.Name, "Command2_Click", "kaydet" 'logging
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+

'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub Command7_Click()
    If Dir(App.path + "\Resimci.exe") <> "" Then
        Shell (App.path & "\Resimci.exe " & App.path & "\Pictures\Pictures" & Anasayfa.dtResim.Recordset.Fields("Resim_Kodu") & ".jpg")
    End If
End Sub
Private Sub Command8_Click()
    Write #7, Time$, Me.Name, "Command8_Click", "Start" 'logging
    Anasayfa.dtSipariþ.Recordset.Edit
    Anasayfa.dtSipariþ.Recordset.Fields("Durum") = "Bitti"
    Anasayfa.dtSipariþ.Recordset.Update
    durumne.Caption = "Durum :    Ýþ Bitti": Command8.Visible = False
    Write #7, Time$, Me.Name, "Command8_Click", "End" 'logging
End Sub
Private Sub Command9_Click()
    Me.Enabled = False: Ödeme.Show
End Sub
Private Sub Command10_Click()
    Write #7, Time$, Me.Name, "Command10_Click", "Start" 'logging
    ParaGiriþi.adý2.Caption = adý.Text
    ParaGiriþi.soyadý2.Caption = soyadý.Text
    ParaGiriþi.Label7.Caption = Anasayfa.dtMüþteri.Recordset.Fields("Musteri_Kodu")
    Me.Enabled = False: ParaGiriþi.Show: ParaGiriþi.startingpoint.Caption = Me.Name
    Write #7, Time$, Me.Name, "Command10_Click", "End" 'logging
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+

Private Sub durumne_Change()
    Write #7, Time$, Me.Name, "durumne_Change", durumne.Caption 'logging
End Sub
Private Sub adý_LostFocus()
    adý.Text = UpperCaseFirstLetter(Trim(adý.Text))
End Sub
Private Sub soyadý_LostFocus()
    soyadý.Text = UpperCaseFirstLetter(Trim(soyadý.Text))
End Sub
Private Sub kumas_Change()
    Dim i As Long: Dim nSel As Long
    If GeriAl = True Or kumas.Text = "" Then GeriAl = False: Exit Sub
    For i = 0 To kumas.ListCount - 1
        If InStr(1, kumas.List(i), kumas.Text, vbTextCompare) = 1 Then
            nSel = kumas.SelStart: kumas.Text = kumas.List(i): kumas.SelStart = nSel: kumas.SelLength = Len(kumas.Text) - nSel
            Exit For
        End If
    Next
End Sub
Private Sub kumas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If kumas.Text <> "" Then GeriAl = True
    End If
End Sub
Private Sub kumas_LostFocus()
    kumas.Text = UpperCaseFirstLetter(Trim(kumas.Text))
End Sub
Private Sub cins_Change()
    Dim i As Long: Dim nSel As Long
    If GeriAl = True Or cins.Text = "" Then GeriAl = False: Exit Sub
    For i = 0 To cins.ListCount - 1
        If InStr(1, cins.List(i), cins.Text, vbTextCompare) = 1 Then
            nSel = cins.SelStart: cins.Text = cins.List(i): cins.SelStart = nSel: cins.SelLength = Len(cins.Text) - nSel
            Exit For
        End If
    Next
End Sub
Private Sub cins_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If cins.Text <> "" Then GeriAl = True
    End If
End Sub
Private Sub cins_LostFocus()
    cins.Text = UpperCaseFirstLetter(Trim(cins.Text))
End Sub
Private Sub fiyat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 13 Then Exit Sub
    If KeyAscii > 57 Or KeyAscii < 47 Then KeyAscii = 0
End Sub
Private Sub tel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 32 Then Exit Sub
    If KeyAscii > 57 Or KeyAscii < 47 Then KeyAscii = 0
End Sub
Private Sub adý_GotFocus(): Call SelectAllText: End Sub
Private Sub soyadý_GotFocus(): Call SelectAllText: End Sub
Private Sub tel_GotFocus(): Call SelectAllText: End Sub
Private Sub fiyat_GotFocus(): Call SelectAllText: End Sub
Private Sub kumas_GotFocus(): Call SelectAllText: End Sub
Private Sub cins_GotFocus(): Call SelectAllText: End Sub
Private Sub bel_GotFocus(): Call SelectAllText: End Sub
Private Sub basen_GotFocus(): Call SelectAllText: End Sub
Private Sub gogus_GotFocus(): Call SelectAllText: End Sub
Private Sub omuz_GotFocus(): Call SelectAllText: End Sub
Private Sub kol_GotFocus(): Call SelectAllText: End Sub
Private Sub boy_GotFocus(): Call SelectAllText: End Sub
Private Sub acik_GotFocus(): Call SelectAllText: End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub Form_Load()
    Write #7, Time$, Me.Name, "Form_Load", "Start" 'logging
    Dim Control As Control: Dim M_ID, S_ID As String: Dim i, j As Integer
    Me.BackColor = rnk_frm_arka: Deðiþtimi = False
    For Each Control In Me
        If TypeOf Control Is OsenXPButton Then Control.BackColor = rnk_btn_arka: Control.BackOver = rnk_btn_arka: Control.ForeColor = rnk_btn_ön: Control.ForeOver = rnk_btn_ön
        If TypeOf Control Is Label Then Control.ForeColor = rnk_frm_ön
        If TypeOf Control Is Frame Then Control.BackColor = rnk_frm_arka: Control.ForeColor = rnk_frm_ön
        If TypeOf Control Is ComboBox Then Control.BackColor = rnk_yazý_arka: Control.ForeColor = rnk_yazý_ön: Control.Locked = True
        If TypeOf Control Is TextBox Then Control.BackColor = rnk_yazý_arka: Control.ForeColor = rnk_yazý_ön: Control.Locked = True
        If TypeOf Control Is DTPicker Then Control.CalendarBackColor = rnk_yazý_arka: Control.CalendarForeColor = rnk_yazý_ön: Control.Enabled = False
        If TypeOf Control Is Line Then Control.BorderColor = rnk_frm_ön
    Next Control
    Open App.path + "\temp2.txt" For Input As #2: Input #2, M_ID: Input #2, S_ID: Close #2
    'kumaþ türleri ve sipariþ türleri combolarý dolduruluyor.
    With Anasayfa.dtKumaþTürü.Recordset 'kumaþ türleri
        .MoveFirst: j = .RecordCount
        For i = 1 To j: kumas.AddItem .Fields("Kumas_Adi"): .MoveNext: Next i
    End With
    With Anasayfa.dtSipariþTürü.Recordset 'sipariþ türleri (cins)
        .MoveFirst: j = .RecordCount
        For i = 1 To j: cins.AddItem .Fields("Siparis_Adi"): .MoveNext: Next i
    End With
    With Anasayfa.dtMüþteri.Recordset
        .MoveFirst
        'müþteri tablosunu baþa sardýktan sonra istediðimiz kayýt gelinceye kadar ileri sarýyoruz.
        While Val(.Fields("Musteri_Kodu")) <> Val(M_ID): .MoveNext: Wend
        'kayýtý bulduk.þimdi o kayýttaki bilgileri forma aktaracaðýz.
        adý.Text = .Fields("Musteri_Adi"): adý2.Caption = adý.Text
        soyadý.Text = .Fields("Musteri_Soyadi"): soyadý2.Caption = soyadý.Text
        tel.Text = .Fields("Musteri_Telefon")
    End With
    'müþteri bilgilerinden sonra sipariþ bilgileri geliyor.
    'yine istediðimiz sipariþe gelinceye kadar ilerletiyoruz.
    With Anasayfa.dtSipariþ.Recordset
        .MoveFirst
        While Val(.Fields("Siparis_Kodu")) <> Val(S_ID): .MoveNext: Wend
        'kayýtý bulduk.þimdi o kayýttaki bilgileri forma aktaracaðýz.
        fiyat.Text = .Fields("Ucret")
        acik.Text = .Fields("Aciklama")
        sip.Value = .Fields("Siparis_Tarihi")
        pro.Value = .Fields("Prova_Tarihi")
        tes.Value = .Fields("Teslim_Tarihi")
        durumne.Caption = "Durum :    Ýþ " + .Fields("Durum"): If .Fields("Durum") = "Bitti" Then Command8.Visible = False
        boy.Text = .Fields("Boy")
        basen.Text = .Fields("Basen")
        bel.Text = .Fields("Bel")
        kol.Text = .Fields("Kol")
        gogus.Text = .Fields("Göðüs")
        omuz.Text = .Fields("Omuz")
    End With
    'sýrada kumaþ türleri
    With Anasayfa.dtKumaþTürü.Recordset
        .MoveFirst
        While Val(.Fields("Kumas_Turu")) <> Val(Anasayfa.dtSipariþ.Recordset.Fields("Kumas_Turu")): .MoveNext: Wend
        kumas2.Caption = .Fields("Kumas_Adi"): kumas.Text = kumas2.Caption: kumas.SelLength = 0
    End With
    'sipariþ türleri
    With Anasayfa.dtSipariþTürü.Recordset
        .MoveFirst
        While Val(.Fields("Siparis_Turleri")) <> Val(Anasayfa.dtSipariþ.Recordset.Fields("Siparis_Turu")): .MoveNext: Wend
        cins2.Caption = .Fields("Siparis_Adi"): cins.Text = cins2.Caption: cins.SelLength = 0
    End With
    'varsa resmi
    With Anasayfa.dtResim.Recordset
        .MoveFirst: j = .RecordCount
        For i = 1 To j
            If Val(.Fields("Siparis_Kodu")) = Val(S_ID) Then Command7.Enabled = True: Exit For
            .MoveNext
        Next i
    End With
    Me.Show: Call frmMain.MDIForm_Resize
    Write #7, Time$, Me.Name, "Form_Load", "Müþteri:" & M_ID & " & Sipariþ:" & S_ID 'logging
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Write #7, Time$, Me.Name, "Form_Unload", "Start" 'logging
    On Local Error Resume Next
    If startingpoint.Caption = "Arama" Then
        Arama.Visible = True
        If Deðiþtimi = True Then Arama.Command2_Click
        Arama.Command1.SetFocus
    ElseIf startingpoint.Caption = "Gsipariþ" Then
        Gsipariþ.Visible = True
        If Deðiþtimi = True Then Gsipariþ.Gör_Click
        Gsipariþ.Command1.SetFocus
    ElseIf startingpoint.Caption = "Maliiþlemler" Then
        Maliiþlemler.Visible = True: Maliiþlemler.MousePointer = 1
        If Deðiþtimi = True Then Maliiþlemler.geri_Click
        Maliiþlemler.geri.SetFocus
    End If
    'logging
    Write #7, Time$, Me.Name, "Form_Unload", "Successful"
End Sub

