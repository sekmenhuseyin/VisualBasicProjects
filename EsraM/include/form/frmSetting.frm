VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form Ayarlar 
   BackColor       =   &H0096E06D&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4695
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6300
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   210
      Top             =   5985
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   2
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   3780
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   6668
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   9887853
      TabCaption(0)   =   "Tema"
      TabPicture(0)   =   "frmSetting.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Görünüm"
      TabPicture(1)   =   "frmSetting.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Yedekleme"
      TabPicture(2)   =   "frmSetting.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Program"
      TabPicture(3)   =   "frmSetting.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Player"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Görünüm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2900
         Left            =   -74750
         TabIndex        =   16
         Top             =   550
         Width           =   5500
         Begin OsenXPCntrl.OsenXPButton renk1 
            Height          =   300
            Left            =   2235
            TabIndex        =   17
            Top             =   1200
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   529
            BTYPE           =   9
            TX              =   "Deðiþtir"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            MICON           =   "frmSetting.frx":007C
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   2
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.OsenXPButton renk2 
            Height          =   300
            Left            =   2235
            TabIndex        =   18
            Top             =   400
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   529
            BTYPE           =   9
            TX              =   "Deðiþtir"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            MICON           =   "frmSetting.frx":0098
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   2
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.OsenXPButton renk3 
            Height          =   300
            Left            =   2235
            TabIndex        =   19
            Top             =   1600
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   529
            BTYPE           =   9
            TX              =   "Deðiþtir"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            MICON           =   "frmSetting.frx":00B4
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   2
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.OsenXPButton renk4 
            Height          =   300
            Left            =   2235
            TabIndex        =   20
            Top             =   2000
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   529
            BTYPE           =   9
            TX              =   "Deðiþtir"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            MICON           =   "frmSetting.frx":00D0
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   2
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.OsenXPButton renk5 
            Height          =   300
            Left            =   2235
            TabIndex        =   21
            Top             =   2400
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   529
            BTYPE           =   9
            TX              =   "Deðiþtir"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            MICON           =   "frmSetting.frx":00EC
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   2
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.OsenXPButton renk6 
            Height          =   300
            Left            =   2235
            TabIndex        =   36
            Top             =   800
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   529
            BTYPE           =   9
            TX              =   "Deðiþtir"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            MICON           =   "frmSetting.frx":0108
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   2
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "From Yazý Rengi"
            Height          =   195
            Left            =   225
            TabIndex        =   38
            Top             =   853
            Width           =   1155
         End
         Begin VB.Label Özel6 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1800
            TabIndex        =   37
            Top             =   800
            Width           =   300
         End
         Begin VB.Label Özel3 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1800
            TabIndex        =   31
            Top             =   1600
            Width           =   300
         End
         Begin VB.Label Özel1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1800
            TabIndex        =   30
            Top             =   1200
            Width           =   300
         End
         Begin VB.Label Özel2 
            BackColor       =   &H0096E06D&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1800
            TabIndex        =   29
            Top             =   400
            Width           =   300
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Form Arka Rengi"
            Height          =   195
            Left            =   225
            TabIndex        =   28
            Top             =   453
            Width           =   1185
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Yazý Alaný Rengi"
            Height          =   195
            Left            =   225
            TabIndex        =   27
            Top             =   1253
            Width           =   1155
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Yazý Rengi"
            Height          =   195
            Left            =   225
            TabIndex        =   26
            Top             =   1653
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Buton Arka Rengi"
            Height          =   195
            Left            =   225
            TabIndex        =   25
            Top             =   2053
            Width           =   1260
         End
         Begin VB.Label özel4 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1800
            TabIndex        =   24
            Top             =   2000
            Width           =   300
         End
         Begin VB.Label özel5 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1800
            TabIndex        =   23
            Top             =   2400
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Buton Yazý Rengi"
            Height          =   195
            Left            =   225
            TabIndex        =   22
            Top             =   2453
            Width           =   1230
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Program Temasý"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2900
         Left            =   250
         TabIndex        =   7
         Top             =   550
         Width           =   5500
         Begin VB.PictureBox Picture1 
            Height          =   1800
            Left            =   250
            ScaleHeight     =   1740
            ScaleWidth      =   1425
            TabIndex        =   12
            Top             =   840
            Width           =   1485
            Begin OsenXPCntrl.OsenXPButton cmd_Tema1 
               Height          =   225
               Left            =   210
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   1050
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   397
               BTYPE           =   3
               TX              =   "AaBb"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
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
               MICON           =   "frmSetting.frx":0124
               UMCOL           =   -1  'True
               SOFT            =   -1  'True
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   2
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.OsenXPButton cmd_Tema2 
               Height          =   225
               Left            =   210
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   1365
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   397
               BTYPE           =   3
               TX              =   "AaBb"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
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
               MICON           =   "frmSetting.frx":0140
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
               BackStyle       =   0  'Transparent
               Caption         =   "EsraM"
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
               Left            =   210
               TabIndex        =   15
               Top             =   630
               Width           =   540
            End
            Begin VB.Image Image1 
               Height          =   330
               Left            =   210
               Stretch         =   -1  'True
               Top             =   210
               Width           =   945
            End
            Begin VB.Image Image2 
               Height          =   1545
               Left            =   105
               Stretch         =   -1  'True
               Top             =   105
               Width           =   1260
            End
         End
         Begin VB.ComboBox Combo 
            Height          =   315
            ItemData        =   "frmSetting.frx":015C
            Left            =   250
            List            =   "frmSetting.frx":0163
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   435
            Width           =   5000
         End
         Begin VB.PictureBox Picture2 
            Height          =   1800
            Left            =   3660
            ScaleHeight     =   1740
            ScaleWidth      =   1530
            TabIndex        =   8
            Top             =   840
            Width           =   1590
            Begin VB.ListBox cmd_Tema5 
               Height          =   1035
               ItemData        =   "frmSetting.frx":0176
               Left            =   105
               List            =   "frmSetting.frx":0180
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   315
               Width           =   1380
            End
            Begin OsenXPCntrl.OsenXPButton cmd_Tema4 
               Height          =   225
               Left            =   840
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   1470
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   397
               BTYPE           =   3
               TX              =   "AaBb"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
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
               MICON           =   "frmSetting.frx":018C
               UMCOL           =   -1  'True
               SOFT            =   -1  'True
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   2
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.OsenXPButton cmd_Tema3 
               Height          =   225
               Left            =   105
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   1470
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   397
               BTYPE           =   3
               TX              =   "AaBb"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   400
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
               MICON           =   "frmSetting.frx":01A8
               UMCOL           =   -1  'True
               SOFT            =   -1  'True
               PICPOS          =   2
               NGREY           =   0   'False
               FX              =   2
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "EsraM"
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
               Left            =   105
               TabIndex        =   32
               Top             =   105
               Width           =   540
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Yedekleme"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2900
         Left            =   -74750
         TabIndex        =   6
         Top             =   550
         Width           =   5500
         Begin OsenXPCntrl.OsenXPButton Command8 
            Height          =   750
            Left            =   45
            TabIndex        =   33
            Top             =   420
            Width           =   5400
            _ExtentX        =   9525
            _ExtentY        =   1323
            BTYPE           =   9
            TX              =   "Veritabanýný Diskete Yedekle"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            MICON           =   "frmSetting.frx":01C4
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
            Height          =   750
            Left            =   45
            TabIndex        =   34
            Top             =   1155
            Width           =   5400
            _ExtentX        =   9525
            _ExtentY        =   1323
            BTYPE           =   9
            TX              =   "Veritabanýný Yedekle"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            MICON           =   "frmSetting.frx":01E0
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   2
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.OsenXPButton Command6 
            Height          =   750
            Left            =   45
            TabIndex        =   35
            Top             =   1890
            Width           =   5400
            _ExtentX        =   9525
            _ExtentY        =   1323
            BTYPE           =   9
            TX              =   "Veritabanýný Yedekten Geri Al"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            MICON           =   "frmSetting.frx":01FC
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
      Begin WMPLibCtl.WindowsMediaPlayer Player 
         Height          =   2505
         Left            =   -74505
         TabIndex        =   5
         Top             =   765
         Width           =   4995
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   43
         autoStart       =   0   'False
         currentMarker   =   0
         invokeURLs      =   0   'False
         baseURL         =   ""
         volume          =   50
         mute            =   -1  'True
         uiMode          =   "none"
         stretchToFit    =   -1  'True
         windowlessVideo =   -1  'True
         enabled         =   0   'False
         enableContextMenu=   0   'False
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   8811
         _cy             =   4419
      End
   End
   Begin OsenXPCntrl.OsenXPButton Command1 
      Default         =   -1  'True
      Height          =   495
      Left            =   1575
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Tamam"
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
      MICON           =   "frmSetting.frx":0218
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
      Cancel          =   -1  'True
      Height          =   495
      Left            =   3135
      TabIndex        =   2
      Top             =   4080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
      MICON           =   "frmSetting.frx":0234
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
      Height          =   495
      Left            =   4725
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Uygula"
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
      MICON           =   "frmSetting.frx":0250
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
Attribute VB_Name = "Ayarlar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Combo_Durum As Boolean
Private Sub Combo_Click()
    If Combo_Durum = False Then Exit Sub
    Write #7, Time$, Me.Name, "Combo_Click", "Start" 'logging
    On Local Error Resume Next
    Dim tmp_rnk_yazý_arka, tmp_rnk_frm_arka, tmp_rnk_frm_ön, tmp_rnk_yazý_ön, tmp_rnk_btn_arka, tmp_rnk_btn_ön, theme_place As String
    If Combo.ListIndex = 0 Then
        cmbindex = 0
        theme_place = Tema_Yeri: tmp_rnk_yazý_arka = rnk_Yazý_Arka: tmp_rnk_frm_arka = rnk_Frm_Arka: tmp_rnk_yazý_ön = rnk_Yazý_Ön: tmp_rnk_btn_arka = rnk_Btn_Arka: tmp_rnk_btn_ön = rnk_Btn_Ön
    Else
        cmbindex = Combo.ListIndex
        theme_place = Themes(Combo.ListIndex - 1).TemaDizin
        tmp_rnk_frm_arka = ReadStringFromIni("Theme", "rnk_frm_arka", "Adsýz", theme_place)
        tmp_rnk_frm_ön = ReadStringFromIni("Theme", "rnk_frm_ön", "Adsýz", theme_place)
        tmp_rnk_yazý_arka = ReadStringFromIni("Theme", "rnk_yazý_arka", "Adsýz", theme_place)
        tmp_rnk_yazý_ön = ReadStringFromIni("Theme", "rnk_yazý_ön", "Adsýz", theme_place)
        tmp_rnk_btn_arka = ReadStringFromIni("Theme", "rnk_btn_arka", "Adsýz", theme_place)
        tmp_rnk_btn_ön = ReadStringFromIni("Theme", "rnk_btn_ön", "Adsýz", theme_place)
    End If
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
'tema önizlemesi
    If Dir(theme_place & "\logo.gif") <> "" Then Image1.Picture = LoadPicture(theme_place & "\logo.gif")
    If Dir(theme_place & "\Background.jpg") <> "" Then Image2.Picture = LoadPicture(theme_place & "\Background.jpg")
    Özel1.BackColor = tmp_rnk_yazý_arka: Özel2.BackColor = tmp_rnk_frm_arka: Özel3.BackColor = tmp_rnk_yazý_ön: özel4.BackColor = tmp_rnk_btn_arka: özel5.BackColor = tmp_rnk_btn_ön: Özel6.BackColor = tmp_rnk_frm_ön
    Label3.ForeColor = tmp_rnk_frm_ön: Label5.ForeColor = tmp_rnk_frm_ön: Picture1.BackColor = tmp_rnk_frm_arka: Picture2.BackColor = tmp_rnk_frm_arka: cmd_Tema5.BackColor = tmp_rnk_yazý_arka: cmd_Tema5.ForeColor = tmp_rnk_yazý_ön
    cmd_Tema1.BackColor = tmp_rnk_btn_arka: cmd_Tema1.BackOver = tmp_rnk_btn_arka: cmd_Tema1.ForeColor = tmp_rnk_btn_ön: cmd_Tema1.ForeOver = tmp_rnk_btn_ön
    cmd_Tema2.BackColor = tmp_rnk_btn_arka: cmd_Tema2.BackOver = tmp_rnk_btn_arka: cmd_Tema2.ForeColor = tmp_rnk_btn_ön: cmd_Tema2.ForeOver = tmp_rnk_btn_ön
    cmd_Tema3.BackColor = tmp_rnk_btn_arka: cmd_Tema3.BackOver = tmp_rnk_btn_arka: cmd_Tema3.ForeColor = tmp_rnk_btn_ön: cmd_Tema3.ForeOver = tmp_rnk_btn_ön
    cmd_Tema4.BackColor = tmp_rnk_btn_arka: cmd_Tema4.BackOver = tmp_rnk_btn_arka: cmd_Tema4.ForeColor = tmp_rnk_btn_ön: cmd_Tema4.ForeOver = tmp_rnk_btn_ön
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*
    Write #7, Time$, Me.Name, "Combo_Click", Combo.List(Combo.ListIndex) 'logging
End Sub
Private Sub Command1_Click()
    Call Command3_Click: Call Command2_Click
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Command3_Click()    'uygula
    Write #7, Time$, Me.Name, "Command3_Click", "Start" 'logging
    On Local Error Resume Next
    Dim i As Integer
    Me.Enabled = False: Me.MousePointer = 11
    rnk_Yazý_Arka = Özel1.BackColor: rnk_Frm_Arka = Özel2.BackColor: rnk_Yazý_Ön = Özel3.BackColor: rnk_Btn_Arka = özel4.BackColor: rnk_Btn_Ön = özel5.BackColor: rnk_Frm_Ön = Özel6.BackColor
    If cmbindex <> 0 Then
        Tema_Adý = Combo.List(cmbindex): Tema_Yeri = Themes(cmbindex - 1).TemaDizin
    End If
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
    Me.BackColor = rnk_Frm_Arka: SSTab.BackColor = rnk_Frm_Arka
    Command1.BackColor = rnk_Btn_Arka: Command1.BackOver = rnk_Btn_Arka: Command1.ForeColor = rnk_Btn_Ön: Command1.ForeOver = rnk_Btn_Ön
    Command2.BackColor = rnk_Btn_Arka: Command2.BackOver = rnk_Btn_Arka: Command2.ForeColor = rnk_Btn_Ön: Command2.ForeOver = rnk_Btn_Ön
    Command3.BackColor = rnk_Btn_Arka: Command3.BackOver = rnk_Btn_Arka: Command3.ForeColor = rnk_Btn_Ön: Command3.ForeOver = rnk_Btn_Ön
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
    Call Anasayfa.AnaGörünüm: Me.Enabled = True: Me.MousePointer = 1: Combo.SetFocus
    Write #7, Time$, Me.Name, "Command3_Click", "uygula" 'logging
End Sub
Private Sub Command5_Click()
    Me.Enabled = False: CopyFile.Caption = "1": CopyFile.Show
End Sub
Private Sub Command6_Click()
    Me.Enabled = False: CopyFile.Caption = "2": CopyFile.Show
End Sub
Private Sub Command8_Click()
    Me.Enabled = False: CopyFile.Caption = "3": CopyFile.Show
End Sub
Private Sub Form_Load()
    Write #7, Time$, Me.Name, "Form_Load", "Start" 'logging
    On Local Error Resume Next
    Dim theme_name As String: Dim i As Integer
    Combo_Durum = False 'ilk açýlýþta combo_click procedürü çalýþmasýn diye...
    Me.Move (frmMain.ScaleWidth - Width) / 2, (frmMain.ScaleHeight - Height) / 2
    'burada varolan tema adlarý comboya ekleniyor.
    For i = 0 To Anasayfa.altTema.Count - 1: Combo.AddItem Themes(i).TemaAd: Next i
    'geçerli tema seçili duruma getiriliyor.
    If cmbindex = 0 Then
        Combo.ListIndex = 0
    Else
        For i = 0 To Combo.ListCount - 1
            If Combo.List(i) = Tema_Adý Then Combo.ListIndex = i: Exit For
        Next i
    End If
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
'formun görünümü
    Me.BackColor = rnk_Frm_Arka: SSTab.BackColor = rnk_Frm_Arka
    Command1.BackColor = rnk_Btn_Arka: Command1.BackOver = rnk_Btn_Arka: Command1.ForeColor = rnk_Btn_Ön: Command1.ForeOver = rnk_Btn_Ön
    Command2.BackColor = rnk_Btn_Arka: Command2.BackOver = rnk_Btn_Arka: Command2.ForeColor = rnk_Btn_Ön: Command2.ForeOver = rnk_Btn_Ön
    Command3.BackColor = rnk_Btn_Arka: Command3.BackOver = rnk_Btn_Arka: Command3.ForeColor = rnk_Btn_Ön: Command3.ForeOver = rnk_Btn_Ön
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
'tema önizlemesi
    Image1.Picture = LoadPicture(Tema_Yeri & "\Logo.gif"): Image2.Picture = LoadPicture(Tema_Yeri & "\Background.jpg")
    Özel1.BackColor = rnk_Yazý_Arka: Özel2.BackColor = rnk_Frm_Arka: Özel3.BackColor = rnk_Yazý_Ön: özel4.BackColor = rnk_Btn_Arka: özel5.BackColor = rnk_Btn_Ön: Özel6.BackColor = rnk_Frm_Ön
    Label3.ForeColor = rnk_Yazý_Ön: Label5.ForeColor = rnk_Yazý_Ön: Picture1.BackColor = rnk_Frm_Arka: Picture2.BackColor = rnk_Frm_Arka: cmd_Tema5.BackColor = rnk_Yazý_Arka: cmd_Tema5.ForeColor = rnk_Yazý_Ön
    cmd_Tema1.BackColor = rnk_Btn_Arka: cmd_Tema1.BackOver = rnk_Btn_Arka: cmd_Tema1.ForeColor = rnk_Btn_Ön: cmd_Tema1.ForeOver = rnk_Btn_Ön
    cmd_Tema2.BackColor = rnk_Btn_Arka: cmd_Tema2.BackOver = rnk_Btn_Arka: cmd_Tema2.ForeColor = rnk_Btn_Ön: cmd_Tema2.ForeOver = rnk_Btn_Ön
    cmd_Tema3.BackColor = rnk_Btn_Arka: cmd_Tema3.BackOver = rnk_Btn_Arka: cmd_Tema3.ForeColor = rnk_Btn_Ön: cmd_Tema3.ForeOver = rnk_Btn_Ön
    cmd_Tema4.BackColor = rnk_Btn_Arka: cmd_Tema4.BackOver = rnk_Btn_Arka: cmd_Tema4.ForeColor = rnk_Btn_Ön: cmd_Tema4.ForeOver = rnk_Btn_Ön
'*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+*+
    Combo_Durum = True: Combo.SetFocus
    Player.URL = App.path + "\Hazýrlayanlar.avi": Player.Controls.play
    Write #7, Time$, Me.Name, "Form_Load", "Successful" 'logging
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Me.Hide: Anasayfa.Enabled = True: Anasayfa.Command1.SetFocus
    Write #7, Time$, Me.Name, "Form_Unload", "Successful" 'logging
End Sub
Private Sub renk1_Click()
    On Local Error GoTo HataControl
    CommonDialog1.ShowColor
    If Val(CommonDialog1.Color) = Val(rnk_Yazý_Arka) Then Exit Sub
    Özel1.BackColor = CommonDialog1.Color: Combo.ListIndex = 0
    cmd_Tema5.BackColor = Özel1.BackColor
    Exit Sub
HataControl:
End Sub
Private Sub renk2_Click()
    On Local Error GoTo HataControl
    CommonDialog1.ShowColor
    If Val(CommonDialog1.Color) = Val(rnk_Frm_Arka) Then Exit Sub
    Özel2.BackColor = CommonDialog1.Color: Combo.ListIndex = 0
    Picture1.BackColor = Özel2.BackColor
    Picture2.BackColor = Özel2.BackColor
    Exit Sub
HataControl:
End Sub
Private Sub renk3_Click()
    On Local Error GoTo HataControl
    CommonDialog1.ShowColor
    If Val(CommonDialog1.Color) = Val(rnk_Yazý_Ön) Then Exit Sub
    Özel3.BackColor = CommonDialog1.Color: Combo.ListIndex = 0
    cmd_Tema5.ForeColor = Özel3.BackColor
    Exit Sub
HataControl:
End Sub
Private Sub renk4_Click()
    On Local Error GoTo HataControl
    CommonDialog1.ShowColor
    If Val(CommonDialog1.Color) = Val(rnk_Btn_Arka) Then Exit Sub
    özel4.BackColor = CommonDialog1.Color: Combo.ListIndex = 0
    cmd_Tema1.BackColor = özel4.BackColor: cmd_Tema1.BackOver = özel4.BackColor
    cmd_Tema2.BackColor = özel4.BackColor: cmd_Tema2.BackOver = özel4.BackColor
    cmd_Tema3.BackColor = özel4.BackColor: cmd_Tema3.BackOver = özel4.BackColor
    cmd_Tema4.BackColor = özel4.BackColor: cmd_Tema4.BackOver = özel4.BackColor
    Exit Sub
HataControl:
End Sub
Private Sub renk5_Click()
    On Local Error GoTo HataControl
    CommonDialog1.ShowColor
    If Val(CommonDialog1.Color) = Val(rnk_Btn_Ön) Then Exit Sub
    özel5.BackColor = CommonDialog1.Color: Combo.ListIndex = 0
    cmd_Tema1.ForeColor = özel5.BackColor: cmd_Tema1.ForeOver = özel5.BackColor
    cmd_Tema2.ForeColor = özel5.BackColor: cmd_Tema2.ForeOver = özel5.BackColor
    cmd_Tema3.ForeColor = özel5.BackColor: cmd_Tema3.ForeOver = özel5.BackColor
    cmd_Tema4.ForeColor = özel5.BackColor: cmd_Tema4.ForeOver = özel5.BackColor
    Exit Sub
HataControl:
End Sub
Private Sub renk6_Click()
    On Local Error GoTo HataControl
    CommonDialog1.ShowColor
    If Val(CommonDialog1.Color) = Val(rnk_Frm_Ön) Then Exit Sub
    Özel6.BackColor = CommonDialog1.Color: Combo.ListIndex = 0
    Label3.ForeColor = Özel3.BackColor
    Label5.ForeColor = Özel3.BackColor
    Exit Sub
HataControl:
End Sub
Private Sub SSTab_Click(PreviousTab As Integer)
    Player.Controls.Stop
    If SSTab.Tab = 0 Then
        Combo.SetFocus 'tema combosu
    ElseIf SSTab.Tab = 1 Then
        Command1.SetFocus 'tamam
    ElseIf SSTab.Tab = 2 Then
        Command8.SetFocus 'diskete yedekle
    ElseIf SSTab.Tab = 3 Then
        Player.Controls.play 'hazýrlayanlar.avi
    End If
End Sub
