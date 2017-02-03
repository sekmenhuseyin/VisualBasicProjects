VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frm_Options 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seçenekler"
   ClientHeight    =   8385
   ClientLeft      =   5475
   ClientTop       =   4515
   ClientWidth     =   11130
   Icon            =   "frm_options.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin OsenXPCntrl.OsenXPButton cmdOK 
      Height          =   405
      Left            =   2910
      TabIndex        =   24
      Top             =   4095
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Tamam"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frm_options.frx":1CFA
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   210
      Top             =   6930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   2
   End
   Begin VB.PictureBox picOptions 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.ListBox List1 
         Height          =   450
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   1170
      End
      Begin VB.Frame fraSample2 
         BackColor       =   &H0080FF80&
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   29
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   31
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   30
         Top             =   675
         Width           =   2055
      End
   End
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   4065
      TabIndex        =   23
      Top             =   4095
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Ýptal"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frm_options.frx":1D16
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdApply 
      Default         =   -1  'True
      Height          =   405
      Left            =   5220
      TabIndex        =   22
      Top             =   4095
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Uygula"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frm_options.frx":1D32
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.TreeView OptionTree 
      Height          =   3900
      Left            =   135
      TabIndex        =   33
      Top             =   113
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   6879
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   1
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   735
      Top             =   6930
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483635
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_options.frx":1D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_options.frx":2628
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_options.frx":4AEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_options.frx":53C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_options.frx":5C9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FR_Content 
      Height          =   3480
      Index           =   2
      Left            =   6405
      TabIndex        =   37
      Tag             =   "Að"
      Top             =   525
      Width           =   4635
      Begin VB.ListBox ListIP 
         Height          =   2010
         ItemData        =   "frm_options.frx":6B78
         Left            =   210
         List            =   "frm_options.frx":6B7A
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   49
         Top             =   735
         Width           =   4215
      End
      Begin OsenXPCntrl.OsenXPButton cmdIPAdd 
         Height          =   300
         Left            =   210
         TabIndex        =   16
         Top             =   2900
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Ekle"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_options.frx":6B7C
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdIPModify 
         Height          =   300
         Left            =   1575
         TabIndex        =   17
         Top             =   2900
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Deðiþtir"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_options.frx":6B98
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdIPDelete 
         Height          =   300
         Left            =   3045
         TabIndex        =   18
         Top             =   2900
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Sil"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_options.frx":6BB4
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdSearchIP 
         Height          =   300
         Left            =   210
         TabIndex        =   15
         Top             =   375
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Otomatik Arama"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_options.frx":6BD0
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame FR_Content 
      Height          =   3480
      Index           =   1
      Left            =   1700
      TabIndex        =   36
      Tag             =   "Görünüm"
      Top             =   4620
      Width           =   4635
      Begin VB.ComboBox ComboTheme 
         Height          =   315
         ItemData        =   "frm_options.frx":6BEC
         Left            =   210
         List            =   "frm_options.frx":6BFF
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2900
         Width           =   4215
      End
      Begin OsenXPCntrl.OsenXPButton cmdBoya 
         Height          =   300
         Index           =   0
         Left            =   2220
         TabIndex        =   10
         Top             =   315
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Deðiþtir"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_options.frx":6C3F
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdBoya 
         Height          =   300
         Index           =   1
         Left            =   2220
         TabIndex        =   11
         Top             =   720
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Deðiþtir"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_options.frx":6C5B
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdBoya 
         Height          =   300
         Index           =   2
         Left            =   2220
         TabIndex        =   12
         Top             =   1140
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Deðiþtir"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_options.frx":6C77
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdBoya 
         Height          =   300
         Index           =   3
         Left            =   2220
         TabIndex        =   13
         Top             =   1567
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Deðiþtir"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_options.frx":6C93
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdBoya 
         Height          =   300
         Index           =   4
         Left            =   2220
         TabIndex        =   14
         Top             =   1987
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Deðiþtir"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_options.frx":6CAF
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   180
         X2              =   4380
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label lbl_Boya 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Arka renk"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   48
         Top             =   375
         Width           =   690
      End
      Begin VB.Label BoyaRengi 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   1815
         TabIndex        =   47
         Top             =   315
         Width           =   300
      End
      Begin VB.Label BoyaRengi 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   4
         Left            =   1815
         TabIndex        =   46
         Top             =   1987
         Width           =   300
      End
      Begin VB.Label lbl_Boya 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Düðme Yazýsý Rengi"
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   45
         Top             =   2040
         Width           =   1425
      End
      Begin VB.Label BoyaRengi 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   3
         Left            =   1815
         TabIndex        =   44
         Top             =   1567
         Width           =   300
      End
      Begin VB.Label lbl_Boya 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Düðme Rengi"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   43
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label BoyaRengi 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   2
         Left            =   1815
         TabIndex        =   42
         Top             =   1155
         Width           =   300
      End
      Begin VB.Label lbl_Boya 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Yazý Rengi"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   41
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label BoyaRengi 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   1815
         TabIndex        =   40
         Top             =   720
         Width           =   300
      End
      Begin VB.Label lbl_Boya 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Yazý Alaný Rengi"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   39
         Top             =   780
         Width           =   1260
      End
   End
   Begin VB.Frame FR_Content 
      Height          =   3480
      Index           =   0
      Left            =   6405
      TabIndex        =   35
      Tag             =   "Uygulama"
      Top             =   4620
      Width           =   4635
      Begin VB.CheckBox CheckAyar 
         Caption         =   "Mesajda Saati Göster"
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   5
         Top             =   2205
         Width           =   2535
      End
      Begin VB.CheckBox CheckAyar 
         Caption         =   "Eski Mesajlarý Sakla"
         Height          =   225
         Index           =   5
         Left            =   210
         TabIndex        =   8
         Top             =   2835
         Width           =   2010
      End
      Begin VB.CheckBox CheckAyar 
         Caption         =   "Her Zaman Üstte"
         Height          =   225
         Index           =   4
         Left            =   210
         TabIndex        =   6
         Top             =   2520
         Width           =   2010
      End
      Begin VB.CheckBox CheckAyar 
         Caption         =   "Simge Durumunda Baþla"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   1
         Top             =   735
         Width           =   2325
      End
      Begin VB.TextBox TextAyar 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   210
         TabIndex        =   4
         Top             =   1530
         Width           =   4110
      End
      Begin VB.CheckBox CheckAyar 
         Caption         =   "Mesaj Uyarý Sesi"
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   2
         Top             =   1095
         Width           =   1905
      End
      Begin VB.CheckBox CheckAyar 
         Caption         =   "Windows Ýle Birlikte Baþla"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   0
         Top             =   375
         Width           =   2745
      End
      Begin OsenXPCntrl.OsenXPButton cmdAyar 
         Height          =   300
         Left            =   2595
         TabIndex        =   3
         Top             =   1095
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Gözat"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_options.frx":6CCB
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdEraseHistory 
         Height          =   300
         Left            =   2595
         TabIndex        =   7
         Top             =   2797
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Geçmiþi Sil"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_options.frx":6CE7
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame FR_Content 
      Height          =   3480
      Index           =   3
      Left            =   1700
      TabIndex        =   38
      Tag             =   "Güvenlik"
      Top             =   543
      Width           =   4635
      Begin VB.TextBox Güven3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   330
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "1"
         Top             =   682
         Width           =   400
      End
      Begin VB.CheckBox Güven2 
         Caption         =   "Yanlýþ Giriþte Bloke Et"
         Height          =   225
         Left            =   225
         TabIndex        =   20
         Top             =   735
         Width           =   1905
      End
      Begin VB.CheckBox Güven0 
         Caption         =   "Açma Parolasý Kullan"
         Height          =   225
         Left            =   225
         TabIndex        =   19
         Top             =   375
         Width           =   1905
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   2730
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   682
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   327681
         OrigLeft        =   2625
         OrigTop         =   840
         OrigRight       =   2880
         OrigBottom      =   1275
         Max             =   11
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdPassword 
         Height          =   300
         Index           =   0
         Left            =   225
         TabIndex        =   50
         Top             =   2415
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Açma Kodunu Deðiþtir"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_options.frx":6D03
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton cmdPassword 
         Height          =   300
         Index           =   1
         Left            =   225
         TabIndex        =   51
         Top             =   2900
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Güvenlik Kodunu Deðiþtir"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483633
         BCOLO           =   -2147483633
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frm_options.frx":6D1F
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   225
         X2              =   4425
         Y1              =   2100
         Y2              =   2100
      End
   End
   Begin VB.Frame FR_Content 
      Height          =   3480
      Index           =   4
      Left            =   1680
      TabIndex        =   52
      Tag             =   "Að"
      Top             =   4620
      Width           =   4635
      Begin WMPLibCtl.WindowsMediaPlayer wmpCredits 
         Height          =   2055
         Left            =   270
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   720
         Width           =   4110
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   100
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   0   'False
         baseURL         =   ""
         volume          =   50
         mute            =   -1  'True
         uiMode          =   "none"
         stretchToFit    =   -1  'True
         windowlessVideo =   0   'False
         enabled         =   0   'False
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   7250
         _cy             =   3625
      End
   End
   Begin VB.Label LBL_Head 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Uygulama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1710
      TabIndex        =   34
      Top             =   113
      Width           =   1395
   End
End
Attribute VB_Name = "frm_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Busy As Boolean
'mesaj uyarý sesi *+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub CheckAyar_Click(index As Integer)
    On Error Resume Next
    If index = 0 Or index = 1 Or index = 3 Or index = 4 Or index = 5 Then Exit Sub
    If CheckAyar(2).Value = 1 Then
        TextAyar.BackColor = vbWhite: TextAyar.Enabled = True: cmdAyar.Enabled = True: cmdAyar.SetFocus
    Else
        cmdAyar.Enabled = False: TextAyar.Enabled = False: TextAyar.BackColor = &HE0E0E0
    End If
End Sub
Private Sub cmdAyar_Click()
    On Local Error GoTo iptalEdildi
    CDialog.Filter = "Ses Dosyalarý|*.mp3;*.wma;*.wav"
    CDialog.InitDir = App.Path & "\sounds\"
    CDialog.ShowOpen
    TextAyar.Text = CDialog.FileName
    Exit Sub
iptalEdildi:
End Sub
Private Sub cmdIPAdd_Click()
    ListIP.AddItem Trim(InputBox("Lütfen geçerli bir IP adresi giriniz !", "Adres ekle"))
    If ListIP.ListCount <> 0 Then ListIP.Selected(0) = True: ListIP.SetFocus
End Sub
Private Sub cmdIPDelete_Click()
    If ListIP.SelCount = 0 Or ListIP.Text <> "" Then
        Dim i, j As Byte: j = ListIP.ListCount
        For i = (j - 1) To 0 Step -1
            If ListIP.Selected(i) = True Then ListIP.RemoveItem i
        Next i
        If ListIP.ListCount <> 0 Then ListIP.Selected(0) = True: ListIP.SetFocus
    End If
End Sub
Private Sub cmdIPModify_Click()
    If ListIP.SelCount = 0 Or ListIP.Text <> "" Then
        ListIP.AddItem Trim(InputBox("Lütfen geçerli bir IP adresi giriniz !", "Adres ekle", ListIP.Text))
        If ListIP.ListCount <> 0 Then ListIP.Selected(0) = True: ListIP.SetFocus
    End If
End Sub
Private Sub cmdPassword_Click(index As Integer)
    frm_Change.Caption = index: frm_Change.Show , Me
End Sub
Private Sub cmdSearchIP_Click()
    Dim l As New cm_LAN: Dim s() As String: Dim i As Integer
    s = Split(l.GetPCList, "||"): ListIP.Clear
    For i = LBound(s) To UBound(s)
        If s(i) <> MyName Then ListIP.AddItem s(i)
    Next i
End Sub
Private Sub ComboTheme_Click()
    If Busy = True Then Exit Sub
    Dim i As Byte
    If ComboTheme.ListIndex = 0 Then
        For i = 0 To 4: BoyaRengi(i).BackColor = Boya(i): Next i
    Else
        For i = 0 To 4: BoyaRengi(i).BackColor = RenkTemalarý(ComboTheme.Text, i): Next i
    End If
End Sub
'mesaj uyarý sesi *+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub TextAyar_GotFocus()
    TextAyar.SelStart = 0: TextAyar.SelLength = Len(TextAyar.Text)
End Sub
Private Sub cmdBoya_Click(index As Integer)
    On Local Error GoTo iptalEdildi
    CDialog.ShowColor
    BoyaRengi(index).BackColor = CDialog.Color
    Exit Sub
iptalEdildi:
End Sub
Private Sub Güven2_Click()
    On Error Resume Next
    If Güven2.Value = 0 Then Güven3.Enabled = False: Güven3.BackColor = &HE0E0E0: UpDown1.Enabled = False Else Güven3.Enabled = True: Güven3.BackColor = &HFFFFFF: UpDown1.Enabled = True: Güven3.SetFocus
End Sub
Private Sub UpDown1_Change()
    If UpDown1.Value = 0 Then UpDown1.Value = 10
    If UpDown1.Value = 11 Then UpDown1.Value = 1
End Sub
Private Sub OptionTree_Click()
    On Error Resume Next
    OptionTree.Nodes(OptionTree.SelectedItem.index).Selected = True
    LBL_Head.Caption = "  " & OptionTree.SelectedItem.Text
    FR_Content(OptionTree.SelectedItem.index - 1).ZOrder vbBringToFront
    Select Case OptionTree.SelectedItem.index
        Case 1: CheckAyar(0).SetFocus: wmpCredits.Controls.stop
        Case 2: ComboTheme.SetFocus: wmpCredits.Controls.stop
        Case 3: cmdSearchIP.SetFocus: wmpCredits.Controls.stop
        Case 4: Güven0.SetFocus: wmpCredits.Controls.stop
        Case 5: cmdCancel.SetFocus: wmpCredits.Controls.play
    End Select
End Sub
Private Sub OptionTree_KeyUp(KeyCode As Integer, Shift As Integer)
    Call OptionTree_Click
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Call cmdApply_Click: Unload Me
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
Private Sub cmdApply_Click()
    On Error Resume Next
    Write #7, Time$, Me.Name, "cmdApply_Click", "Start" 'logging
    Dim i As Byte
    '[Uygulama]
    For i = 0 To 2: Ayar(i) = CheckAyar(i).Value: Next i: Ayar(3) = TextAyar.Text: For i = 4 To 6: Ayar(i) = CheckAyar(i - 1).Value: Next i
    If CheckAyar(0).Value = 1 Then
        SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "iso MSN", App.Path & "\" & App.EXEName & ".exe"
    Else
        DelSetting HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "iso MSN"
    End If
    '[Görünüm]
    For i = 0 To 4: Boya(i) = BoyaRengi(i).BackColor: Next i
    BoyaTheme = ComboTheme.List(ComboTheme.ListIndex)
    '[Að]
    If ListIP.ListCount > 0 Then
        For i = 0 To ListIP.ListCount
            NetName(i) = ListIP.List(i): NetPath(i) = "\\" & NetName(i) & "\c\sohbet.txt"
        Next i
    End If
    For i = ListIP.ListCount To 254: NetPath(i) = "": NetName(i) = "": Next i
    '[Güvenlik]
    Güven(0) = Güven0.Value
    Güven(2) = Güven2.Value
    Güven(3) = Güven3.Text
    'others
    Call frm_Messenger.AnaGörünüm
    Write #7, Time$, Me.Name, "cmdApply_Click", "Successful" 'logging
End Sub
Private Sub Form_Load()
    Dim i As Integer: Busy = True
    Write #7, Time$, Me.Name, "Form_Load", "Start" 'logging
    '[Uygulama]
    CheckAyar(0).Value = Ayar(0)
    CheckAyar(1).Value = Ayar(1)
    CheckAyar(2).Value = Ayar(2)
    TextAyar.Text = Ayar(3)
    CheckAyar(3).Value = Ayar(4)
    CheckAyar(4).Value = Ayar(5)
    CheckAyar(5).Value = Ayar(6)
    '[Görünüm]
    BoyaRengi(0).BackColor = Boya(0)
    BoyaRengi(1).BackColor = Boya(1)
    BoyaRengi(2).BackColor = Boya(2)
    BoyaRengi(3).BackColor = Boya(3)
    BoyaRengi(4).BackColor = Boya(4)
    For i = 0 To ComboTheme.ListCount
        If ComboTheme.List(i) = BoyaTheme Then ComboTheme.ListIndex = i: Exit For
    Next i
    '[Güvenlik]
    Güven0.Value = Güven(0)
    Güven2.Value = Güven(2)
    Güven3.Text = Güven(3)
    '[Að]
    For i = 0 To 254
        If NetName(i) <> "" Then ListIP.AddItem NetName(i) Else Exit For
    Next i
    If ListIP.ListCount <> 0 Then ListIP.Selected(0) = True
    'option tree & others
    OptionTree.Nodes.Clear
    OptionTree.Nodes.Add , , "Uygulama", "Uygulama", 1
    OptionTree.Nodes.Add , , "Görünüm", "Görünüm ", 2
    OptionTree.Nodes.Add , , "Að", "Að          ", 3
    OptionTree.Nodes.Add , , "Güvenlik", "Güvenlik ", 4
    OptionTree.Nodes.Add , , "Hakkýnda", "Hakkýnda", 5
    OptionTree.Nodes.Item(1).Selected = True: Call OptionTree_Click
    Me.Width = 6600: Me.Height = 5100: For i = 0 To 4: FR_Content(i).Move 1700, 550: Next i
    Busy = False: wmpCredits.URL = App.Path & "\sounds\credits.avi"
    If Ayar(5) = "1" Then AlwaysOnTop Me, True
    Write #7, Time$, Me.Name, "Form_Load", "Successful" 'logging
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frm_Messenger.ToolMSN.Highlight "btnOptions", 0:
    frm_Messenger.Enabled = True: frm_Messenger.Text_Giden.SetFocus: Unload Me
    Write #7, Time$, Me.Name, "Form_Unload", "Successful" 'logging
End Sub
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+

