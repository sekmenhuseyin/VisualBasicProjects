VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00CDB75F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Malzeme Girisi - Money"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   Enabled         =   0   'False
   Icon            =   "malzemegirisi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "malzemegirisi.frx":030A
   ScaleHeight     =   8145
   ScaleWidth      =   11910
   Begin VB.CommandButton Cmmd7 
      Caption         =   "��k�s"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   315
      TabIndex        =   49
      Top             =   4590
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd5 
      Caption         =   "Fiyat Listesi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   315
      TabIndex        =   48
      Top             =   3150
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd4 
      Caption         =   "Stok Kontrol�"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   315
      TabIndex        =   47
      Top             =   2415
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd3 
      Caption         =   "Malzeme Listesi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   315
      TabIndex        =   46
      Top             =   1680
      Width           =   2985
   End
   Begin VB.CommandButton Cmmd2 
      Caption         =   "Malzeme Sat�s�"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   315
      TabIndex        =   45
      Top             =   945
      Width           =   2985
   End
   Begin VB.ListBox ListS�rala 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   645
      ItemData        =   "malzemegirisi.frx":1C8C
      Left            =   7875
      List            =   "malzemegirisi.frx":1C8E
      Sorted          =   -1  'True
      TabIndex        =   44
      Top             =   525
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   8480
      Picture         =   "malzemegirisi.frx":1C90
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   43
      Top             =   7245
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5800
      Picture         =   "malzemegirisi.frx":3212
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   42
      Top             =   7350
      Width           =   240
   End
   Begin VB.CommandButton CmdReset 
      Caption         =   "S�f�rla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   8400
      MaskColor       =   &H8000000F&
      TabIndex        =   12
      Top             =   7140
      Width           =   2000
   End
   Begin VB.CommandButton CmdKaydet 
      Caption         =   "Kaydet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   5670
      MaskColor       =   &H8000000F&
      TabIndex        =   11
      Top             =   7140
      Width           =   2000
   End
   Begin VB.CommandButton Command12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8715
      Picture         =   "malzemegirisi.frx":55B4
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Se�ili Garanti S�resini De�istir"
      Top             =   3200
      Width           =   435
   End
   Begin VB.CommandButton Command11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8715
      Picture         =   "malzemegirisi.frx":58BE
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Se�ili T�r�n Ad� De�istir"
      Top             =   2600
      Width           =   435
   End
   Begin VB.CommandButton Command10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8715
      Picture         =   "malzemegirisi.frx":5BC8
      Style           =   1  'Graphical
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "Se�ili Modelin Ad� De�istir"
      Top             =   2000
      Width           =   435
   End
   Begin VB.CommandButton Command9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8715
      Picture         =   "malzemegirisi.frx":5ED2
      Style           =   1  'Graphical
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   "Se�ili Markan�n Ad� De�istir"
      Top             =   1395
      Width           =   435
   End
   Begin VB.CommandButton Command8 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8295
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "Se�ili Garanti S�resini Kald�r"
      Top             =   3200
      Width           =   330
   End
   Begin VB.CommandButton Command7 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7875
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "Garanti S�resi Ekle"
      Top             =   3200
      Width           =   330
   End
   Begin VB.CommandButton Command6 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8295
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Se�ili T�r� Kald�r"
      Top             =   2600
      Width           =   330
   End
   Begin VB.CommandButton Command5 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7875
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "T�r Ekle"
      Top             =   2600
      Width           =   330
   End
   Begin VB.CommandButton Command4 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8295
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Se�ili Modeli Kald�r"
      Top             =   2000
      Width           =   330
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7875
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Model Ekle"
      Top             =   2000
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8295
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Se�ili Markay� Kald�r"
      Top             =   1395
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7875
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Marka Ekle"
      Top             =   1395
      Width           =   330
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "malzemegirisi.frx":61DC
      Left            =   5085
      List            =   "malzemegirisi.frx":61E3
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3200
      Width           =   2700
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "malzemegirisi.frx":61EC
      Left            =   5085
      List            =   "malzemegirisi.frx":61F3
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2600
      Width           =   2700
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "malzemegirisi.frx":6205
      Left            =   5085
      List            =   "malzemegirisi.frx":620C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2000
      Width           =   2700
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "malzemegirisi.frx":621E
      Left            =   5085
      List            =   "malzemegirisi.frx":6225
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1395
      Width           =   2700
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   5085
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   3800
      Width           =   6690
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9435
      TabIndex        =   10
      Top             =   6090
      Width           =   2340
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9435
      TabIndex        =   9
      Top             =   5400
      Width           =   2340
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9435
      TabIndex        =   6
      Top             =   2000
      Width           =   2340
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9435
      TabIndex        =   7
      Top             =   3200
      Width           =   2340
   End
   Begin VB.TextBox TxtTarih 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5085
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5400
      Width           =   2000
   End
   Begin VB.TextBox TxtSaat 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5085
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6090
      Width           =   2000
   End
   Begin VB.CommandButton Cmmd1 
      Caption         =   "Malzeme Girisi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   315
      TabIndex        =   0
      Top             =   210
      Width           =   2985
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   6300
   End
   Begin VB.CommandButton Cmmd6 
      Caption         =   "Hakk�nda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   315
      TabIndex        =   1
      Top             =   3885
      Width           =   2985
   End
   Begin VB.Line Line2 
      X1              =   9345
      X2              =   11775
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Line Line1 
      X1              =   9345
      X2              =   9345
      Y1              =   1365
      Y2              =   3675
   End
   Begin VB.Label Modeli 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modeli"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   3885
      TabIndex        =   29
      Top             =   2000
      Width           =   795
   End
   Begin VB.Label Markas� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Markas�"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3885
      TabIndex        =   28
      Top             =   1400
      Width           =   945
   End
   Begin VB.Label Cinsi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T�r�"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   3900
      TabIndex        =   27
      Top             =   2595
      Width           =   555
   End
   Begin VB.Label Serino 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seri \ Barkod No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   7350
      TabIndex        =   26
      Top             =   6195
      Width           =   1980
   End
   Begin VB.Label �zellikler 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�zellikler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3900
      TabIndex        =   25
      Top             =   3800
      Width           =   1110
   End
   Begin VB.Label Garanti 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Garanti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   3900
      TabIndex        =   23
      Top             =   3200
      Width           =   915
   End
   Begin VB.Label YeniMalzemeGirisi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Malzeme Girisi"
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
      Left            =   3780
      TabIndex        =   22
      Top             =   105
      Width           =   2070
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fatura Tarihi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9450
      TabIndex        =   21
      Top             =   1400
      Width           =   1545
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fatura Numaras�"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9450
      TabIndex        =   20
      Top             =   2600
      Width           =   2010
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tarih"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3900
      TabIndex        =   19
      Top             =   5505
      Width           =   615
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3900
      TabIndex        =   18
      Top             =   6195
      Width           =   585
   End
   Begin VB.Line Line9 
      BorderWidth     =   3
      X1              =   3780
      X2              =   5880
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label LabelKontrol 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   1365
      TabIndex        =   15
      Top             =   6405
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Labelyaz� 
      BackStyle       =   0  'Transparent
      Caption         =   "Bunlar� Biliyormuydunuz;"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   200
      TabIndex        =   14
      Top             =   5450
      Width           =   3015
   End
   Begin VB.Label lblipucu 
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
      ForeColor       =   &H00808000&
      Height          =   2600
      Left            =   100
      TabIndex        =   13
      Top             =   5750
      Width           =   3500
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "malzemegirisi.frx":6237
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3750
   End
   Begin VB.Label Fiyat� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fiyat�"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   7350
      TabIndex        =   24
      Top             =   5505
      Width           =   660
   End
   Begin VB.Image Background 
      Height          =   495
      Left            =   10560
      Picture         =   "malzemegirisi.frx":2C7D2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim YeniMarka, YeniModel, YeniT�r, YeniGaranti, Soru, Newmkt, Kontrol
Private Sub Cmmd2_Click()
Form2.Show: Me.Hide
End Sub
Private Sub Cmmd3_Click()
Form3.Show: Me.Hide
End Sub
Private Sub Cmmd4_Click()
Form4.Show: Me.Hide
End Sub
Private Sub Cmmd5_Click()
Form5.Show: Me.Hide
End Sub
Private Sub Cmmd7_Click()
End
End Sub
Private Sub cmdKaydet_Click()
On Error GoTo Hata
'-------------------------kontrol-------------------------
Kontrol = 0
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
KontrolSorusu = MsgBox("Eksik bilgi girdiniz. Devam etmek istiyormusunuz?", vbYesNo)
If KontrolSorusu = vbNo Then Exit Sub
End If
'**************************malzeme listesi i�in bilgi girisi*************************
Open App.Path + "\stuff\stok.mlz" For Input As #1
bas1:
If EOF(1) Then GoTo son1
'input olarak mrk(marka), mdl(model), cns(cins), grn(garanti), mkt(miktar)
Input #1, mrk, mdl, cns, grn, mkt
If mrk = Combo1.Text And mdl = Combo2.Text And cns = Combo3.Text And grn = Combo4.Text Then
'bulursa kontrol 1 oluyor ve miktar� bir artt�r�yor, eskiye ek b�l�m�ne g�nderilecek
Kontrol = 1: Newmkt = mkt + 1: GoTo son1
Else
GoTo bas1
End If
son1:
Close #1
'e�er daha �nceden b�yle bir malzeme varsa miktar� artt�racak
'yoksa yeni malzeme girisi yap�lacak
If Kontrol = 0 Then GoTo YeniKay�t
If Kontrol = 1 Then GoTo EskiyeDevam
'-------------------------yenikay�t-------------------------
YeniKay�t:
istek = MsgBox("Belirtti�iniz �zellikte bir malzeme kay�tl� de�il. �imdi kaydetmek istermisiniz?", vbYesNo)
If istek = vbNo Then Exit Sub
Open App.Path + "\stuff\stok.mlz" For Append As #1
Write #1, Combo1.Text, Combo2.Text, Combo3.Text, Combo4.Text, "1"
Close #1
GoTo Kay�t�ndevam�
'-------------------------eskiye ek-------------------------
EskiyeDevam:
Open App.Path + "\stuff\stok.mlz" For Input As #1
Open App.Path + "\stuff\temp.mlz" For Output As #2
'd�ng� baslar, e�er ger�ek dosyadaki isim se�ili isim de�ilse temp e yaz
'aksi takdirde yeni de�erleri yaz ve di�erine ge�
bas2:
If EOF(1) Then GoTo son2
Input #1, mrk, mdl, cns, grn, mkt
If mrk = Combo1.Text And mdl = Combo2.Text And cns = Combo3.Text And grn = Combo4.Text Then
'kaydedilen dosya ise
Write #2, mrk, mdl, cns, grn, Newmkt
Else
'de�il ise
Write #2, mrk, mdl, cns, grn, mkt
End If
GoTo bas2
son2:
't�m dosyalar� kapat�r.
'ger�ek dosyay� siler ve ge�iciyi ger�ek yapar
Close #1
Close #2
Kill App.Path + "\stuff\stok.mlz"
Name App.Path + "\stuff\temp.mlz" As App.Path + "\stuff\stok.mlz"
'******************toplam malzeme bilgilerini (malzeme sat�s� i�in) girer*****************
Kay�t�ndevam�:
Open App.Path + "\stuff\market.mlz" For Append As #1
Write #1, Combo1.Text, Combo2.Text, Combo3.Text, Combo4.Text, Text1, Text2, Text3, Text4, Text5, TxtTarih, TxtSaat
Close #1
'******************stok kontrol� i�in -hem giris hem ��k�s-*******************************
'sadece belli basl� bilgiler yaz�lacak...
Open App.Path + "\stuff\hst.mlz" For Append As #1
Write #1, Combo1.Text, Combo2.Text, Combo3.Text, Text3, Text4, Text5, TxtTarih, "Al�nd�"
Close #1
've bilgiler sa�sa�lim yaz�ld���na dair mesaj verilir
MsgBox "   Bilgileriniz Kaydedildi   "
've temizlik
Form_Activate
Exit Sub
'ne olur olmaz diye yapt�m...
Hata:
MsgBox "   Bilinmeyen bir nedenden dolay� bilgileriniz kaydedilmedi   "
End Sub
Private Sub CmdReset_Click()
Form_Activate
End Sub
Private Sub Command1_Click()
ListS�rala.Clear
'eklenecek maddenin ad�n� sorar
YeniMarka = InputBox("L�tfen yeni markan�n ad�n� yaz�n�z", "Yeni Marka")
'e�er madde ismi bos ise hataya g�nderir
If Len(YeniMarka) = 0 Or YeniMarka = String$(Len(YeniMarka), " ") Then GoTo Kay�tHatas�
'eklenecek maddenin dosyas�n� a�ar
Open App.Path + "\data\cmd1.nfo" For Input As #1
Bas:
'dosyadaki t�m isimleri listeye ekler, yeni ismi de ekler
If EOF(1) Then GoTo Son
Input #1, mark
ListS�rala.AddItem mark
GoTo Bas
Son:
Close #1
ListS�rala.AddItem UCase(YeniMarka)
ListS�rala.ListIndex = 0
'sonra listedeki maddeleri dosyaya yazar
Open App.Path + "\data\cmd1.nfo" For Output As #1
ilk:
'yeni maddeyi ekler -listede se�ili olan�-
Write #1, ListS�rala.List(ListS�rala.ListIndex)
'e�er listenin sonuncusu se�iliyse bitir
If ListS�rala.ListIndex = ListS�rala.ListCount - 1 Then GoTo iki
'listenin bir alt�na ge�
ListS�rala.ListIndex = ListS�rala.ListIndex + 1
GoTo ilk
iki:
'dosyay� kapat�r
Close #1
'birlesik kutuya yazar
MarkaYenile
Exit Sub
'iste hata b�l�m�
Kay�tHatas�:
MsgBox "Markan�n ismi do�ru yaz�lmad���ndan kay�t ger�eklestirelemedi"
End Sub
Private Sub Command2_Click()
On Error GoTo Hata
'e�er combonun ilk maddesi se�ili (yani "<Bilinmeyen>") ise hata yapt�r.
If Combo1.ListIndex = 0 Then GoTo Hata
Soru = MsgBox("'" + Combo1.Text + "'" + " adl� markay� silmek istiyormusunuz?", vbYesNo)
If Soru = vbNo Then Exit Sub
'bi ger�ek dosya bi de temp a��l�r.
Open App.Path + "\data\cmd1.nfo" For Input As #1
Open App.Path + "\data\cmd2.nfo" For Output As #2
'd�ng� baslar, e�er ger�ek dosyadaki isim se�ili isim de�ilse temp e yaz
'aksi takdirde yazma ve di�erine ge�
Bas:
If EOF(1) Then GoTo Son
Input #1, YeniMarka
If Combo1.Text <> YeniMarka Then
Write #2, YeniMarka
GoTo Bas
Else
GoTo Bas
End If
Son:
't�m dosyalar� kapat�r.
'ger�ek dosyay� siler ve ge�iciyi ger�ek yapar
Close #1
Close #2
Kill App.Path + "\data\cmd1.nfo"
Name App.Path + "\data\cmd2.nfo" As App.Path + "\data\cmd1.nfo"
'comboyu yeniler
MarkaYenile
Exit Sub
Hata:
MsgBox "�stedi�iniz islem ger�eklestirilmedi."
End Sub
Private Sub Command3_Click()
ListS�rala.Clear
'eklenecek maddenin ad�n� sorar
YeniModel = InputBox("L�tfen yeni modelin ad�n� yaz�n�z", "Yeni Model")
'e�er madde ismi bos ise hataya g�nderir
If Len(YeniModel) = 0 Or YeniModel = String$(Len(YeniModel), " ") Then GoTo Kay�tHatas�
'eklenecek maddenin dosyas�n� a�ar
Open App.Path + "\data\cmd3.nfo" For Input As #1
Bas:
'dosyadaki t�m isimleri listeye ekler, yeni ismi de ekler
If EOF(1) Then GoTo Son
Input #1, model
ListS�rala.AddItem model
GoTo Bas
Son:
Close #1
ListS�rala.AddItem UCase(YeniModel)
ListS�rala.ListIndex = 0
'sonra listedeki maddeleri dosyaya yazar
Open App.Path + "\data\cmd3.nfo" For Output As #1
ilk:
'yeni maddeyi ekler -listede se�ili olan�-
Write #1, ListS�rala.List(ListS�rala.ListIndex)
'e�er listenin sonuncusu se�iliyse bitir
If ListS�rala.ListIndex = ListS�rala.ListCount - 1 Then GoTo iki
'listenin bir alt�na ge�
ListS�rala.ListIndex = ListS�rala.ListIndex + 1
GoTo ilk
iki:
'dosyay� kapat�r
Close #1
'birlesik kutuya yazar
ModelYenile
Exit Sub
'iste hata b�l�m�
Kay�tHatas�:
MsgBox "Modelin ismi do�ru yaz�lmad���ndan kay�t ger�eklestirelemedi"
End Sub
Private Sub Command4_Click()
On Error GoTo Hata
'e�er combonun ilk maddesi se�ili (yani "<Bilinmeyen>") ise hata yapt�r.
If Combo2.ListIndex = 0 Then GoTo Hata
Soru = MsgBox("'" + Combo2.Text + "'" + " adl� modeli silmek istiyormusunuz?", vbYesNo)
If Soru = vbNo Then Exit Sub
'bi ger�ek dosya bi de temp a��l�r.
Open App.Path + "\data\cmd3.nfo" For Input As #1
Open App.Path + "\data\cmd4.nfo" For Output As #2
'd�ng� baslar, e�er ger�ek dosyadaki isim se�ili isim de�ilse temp e yaz
'aksi takdirde yazma ve di�erine ge�
Bas:
If EOF(1) Then GoTo Son
Input #1, YeniModel
If Combo2.Text <> YeniModel Then
Write #2, YeniModel
GoTo Bas
Else
GoTo Bas
End If
Son:
't�m dosyalar� kapat�r.
'ger�ek dosyay� siler ve ge�iciyi ger�ek yapar
Close #1
Close #2
Kill App.Path + "\data\cmd3.nfo"
Name App.Path + "\data\cmd4.nfo" As App.Path + "\data\cmd3.nfo"
'comboyu yeniler
ModelYenile
Exit Sub
Hata:
MsgBox "�stedi�iniz islem ger�eklestirilmedi."
End Sub
Private Sub Command5_Click()
ListS�rala.Clear
'eklenecek maddenin ad�n� sorar
YeniT�r = InputBox("L�tfen yeni t�r�n ad�n� yaz�n�z", "Yeni T�r")
'e�er madde ismi bos ise hataya g�nderir
If Len(YeniT�r) = 0 Or YeniT�r = String$(Len(YeniT�r), " ") Then GoTo Kay�tHatas�
'eklenecek maddenin dosyas�n� a�ar
Open App.Path + "\data\cmd5.nfo" For Input As #1
Bas:
'dosyadaki t�m isimleri listeye ekler, yeni ismi de ekler
If EOF(1) Then GoTo Son
Input #1, t�r
ListS�rala.AddItem t�r
GoTo Bas
Son:
Close #1
ListS�rala.AddItem UCase(YeniT�r)
ListS�rala.ListIndex = 0
'sonra listedeki maddeleri dosyaya yazar
Open App.Path + "\data\cmd5.nfo" For Output As #1
ilk:
'yeni maddeyi ekler -listede se�ili olan�-
Write #1, ListS�rala.List(ListS�rala.ListIndex)
'e�er listenin sonuncusu se�iliyse bitir
If ListS�rala.ListIndex = ListS�rala.ListCount - 1 Then GoTo iki
'listenin bir alt�na ge�
ListS�rala.ListIndex = ListS�rala.ListIndex + 1
GoTo ilk
iki:
'dosyay� kapat�r
Close #1
'birlesik kutuya yazar
T�rYenile
Exit Sub
'iste hata b�l�m�
Kay�tHatas�:
MsgBox "T�r�n ismi do�ru yaz�lmad���ndan kay�t ger�eklestirelemedi"
End Sub
Private Sub Command6_Click()
On Error GoTo Hata
'e�er combonun ilk maddesi se�ili (yani "<Bilinmeyen>") ise hata yapt�r.
If Combo3.ListIndex = 0 Then GoTo Hata
Soru = MsgBox("'" + Combo3.Text + "'" + " adl� t�r� silmek istiyormusunuz?", vbYesNo)
If Soru = vbNo Then Exit Sub
'bi ger�ek dosya bi de temp a��l�r.
Open App.Path + "\data\cmd5.nfo" For Input As #1
Open App.Path + "\data\cmd6.nfo" For Output As #2
'd�ng� baslar, e�er ger�ek dosyadaki isim se�ili isim de�ilse temp e yaz
'aksi takdirde yazma ve di�erine ge�
Bas:
If EOF(1) Then GoTo Son
Input #1, YeniT�r
If Combo3.Text <> YeniT�r Then
Write #2, YeniT�r
GoTo Bas
Else
GoTo Bas
End If
Son:
't�m dosyalar� kapat�r.
'ger�ek dosyay� siler ve ge�iciyi ger�ek yapar
Close #1
Close #2
Kill App.Path + "\data\cmd5.nfo"
Name App.Path + "\data\cmd6.nfo" As App.Path + "\data\cmd5.nfo"
'comboyu yeniler
T�rYenile
Exit Sub
Hata:
MsgBox "�stedi�iniz islem ger�eklestirilmedi."
End Sub
Private Sub Command7_Click()
ListS�rala.Clear
'eklenecek maddenin ad�n� sorar
YeniGaranti = InputBox("L�tfen yeni garanti s�resini yaz�n�z", "Yeni Garanti")
'e�er madde ismi bos ise hataya g�nderir
If Len(YeniGaranti) = 0 Or YeniGaranti = String$(Len(YeniGaranti), " ") Then GoTo Kay�tHatas�
'eklenecek maddenin dosyas�n� a�ar
Open App.Path + "\data\cmd7.nfo" For Input As #1
Bas:
'dosyadaki t�m isimleri listeye ekler, yeni ismi de ekler
If EOF(1) Then GoTo Son
Input #1, gara
ListS�rala.AddItem gara
GoTo Bas
Son:
Close #1
ListS�rala.AddItem UCase(YeniGaranti)
ListS�rala.ListIndex = 0
'sonra listedeki maddeleri dosyaya yazar
Open App.Path + "\data\cmd7.nfo" For Output As #1
ilk:
'yeni maddeyi ekler -listede se�ili olan�-
Write #1, ListS�rala.List(ListS�rala.ListIndex)
'e�er listenin sonuncusu se�iliyse bitir
If ListS�rala.ListIndex = ListS�rala.ListCount - 1 Then GoTo iki
'listenin bir alt�na ge�
ListS�rala.ListIndex = ListS�rala.ListIndex + 1
GoTo ilk
iki:
'dosyay� kapat�r
Close #1
'birlesik kutuya yazar
GarantiYenile
Exit Sub
'iste hata b�l�m�
Kay�tHatas�:
MsgBox "Garantinin s�resi do�ru yaz�lmad���ndan kay�t ger�eklestirelemedi"
End Sub
Private Sub Command8_Click()
On Error GoTo Hata
'e�er combonun ilk maddesi se�ili (yani "<Bilinmeyen>") ise hata yapt�r.
If Combo4.ListIndex = 0 Then GoTo Hata
Soru = MsgBox("'" + Combo4.Text + "'" + " s�reli garantiyi silmek istiyormusunuz?", vbYesNo)
If Soru = vbNo Then Exit Sub
'bi ger�ek dosya bi de temp a��l�r.
Open App.Path + "\data\cmd7.nfo" For Input As #1
Open App.Path + "\data\cmd8.nfo" For Output As #2
'd�ng� baslar, e�er ger�ek dosyadaki isim se�ili isim de�ilse temp e yaz
'aksi takdirde yazma ve di�erine ge�
Bas:
If EOF(1) Then GoTo Son
Input #1, YeniGaranti
If Combo4.Text <> YeniGaranti Then
Write #2, YeniGaranti
GoTo Bas
Else
GoTo Bas
End If
Son:
't�m dosyalar� kapat�r.
'ger�ek dosyay� siler ve ge�iciyi ger�ek yapar
Close #1
Close #2
Kill App.Path + "\data\cmd7.nfo"
Name App.Path + "\data\cmd8.nfo" As App.Path + "\data\cmd7.nfo"
'comboyu yeniler
GarantiYenile
Exit Sub
Hata:
MsgBox "�stedi�iniz islem ger�eklestirilmedi."
End Sub
Private Sub Command9_Click()
ListS�rala.Clear
On Error GoTo Hata
'e�er combonun ilk maddesi se�ili (yani "<Bilinmeyen>") ise hata yapt�r.
If Combo1.ListIndex = 0 Then GoTo Hata
Soru = InputBox(Combo1.Text + " adl� markan�n yeni ad�n� giriniz.", "Marka ad� de�istir")
If Len(Soru) = 0 Or Soru = String$(Len(Soru), " ") Then GoTo Hata
'bi ger�ek dosya bi de temp a��l�r.
Open App.Path + "\data\cmd1.nfo" For Input As #1
'd�ng� baslar, e�er ger�ek dosyadaki isim se�ili isim de�ilse temp e yaz
'aksi takdirde yazma ve di�erine ge�
ilk:
'dosyadaki t�m isimleri(de�isecek olan hari�) listeye ekler
'yeni ismi de ekler
If EOF(1) Then GoTo iki
Input #1, mark
If Combo1.Text <> mark Then
ListS�rala.AddItem mark
GoTo ilk
Else
GoTo ilk
End If
iki:
Close #1
ListS�rala.AddItem UCase(Soru)
ListS�rala.ListIndex = 0
'sonra listedeki maddeleri dosyaya yazar
Open App.Path + "\data\cmd1.nfo" For Output As #1
Bas:
'yeni maddeyi ekler -listede se�ili olan�-
Write #1, ListS�rala.List(ListS�rala.ListIndex)
'e�er listenin sonuncusu se�iliyse bitir
If ListS�rala.ListIndex = ListS�rala.ListCount - 1 Then GoTo Son
'listenin bir alt�na ge�
ListS�rala.ListIndex = ListS�rala.ListIndex + 1
GoTo Bas
Son:
't�m dosyalar� kapat�r.
Close #1
'comboyu yeniler
MarkaYenile
Exit Sub
Hata:
MsgBox "�stedi�iniz islem ger�eklestirilmedi."
End Sub
Private Sub Command10_Click()
ListS�rala.Clear
On Error GoTo Hata
'e�er combonun ilk maddesi se�ili (yani "<Bilinmeyen>") ise hata yapt�r.
If Combo2.ListIndex = 0 Then GoTo Hata
Soru = InputBox(Combo2.Text + " adl� modelin yeni ad�n� giriniz.", "Model ad� de�istir")
If Len(Soru) = 0 Or Soru = String$(Len(Soru), " ") Then GoTo Hata
'bi ger�ek dosya bi de temp a��l�r.
Open App.Path + "\data\cmd3.nfo" For Input As #1
'd�ng� baslar, e�er ger�ek dosyadaki isim se�ili isim de�ilse temp e yaz
'aksi takdirde yazma ve di�erine ge�
ilk:
'dosyadaki t�m isimleri(de�isecek olan hari�) listeye ekler
'yeni ismi de ekler
If EOF(1) Then GoTo iki
Input #1, mode
If Combo2.Text <> mode Then
ListS�rala.AddItem mode
GoTo ilk
Else
GoTo ilk
End If
iki:
Close #1
ListS�rala.AddItem UCase(Soru)
ListS�rala.ListIndex = 0
'sonra listedeki maddeleri dosyaya yazar
Open App.Path + "\data\cmd3.nfo" For Output As #1
Bas:
'yeni maddeyi ekler -listede se�ili olan�-
Write #1, ListS�rala.List(ListS�rala.ListIndex)
'e�er listenin sonuncusu se�iliyse bitir
If ListS�rala.ListIndex = ListS�rala.ListCount - 1 Then GoTo Son
'listenin bir alt�na ge�
ListS�rala.ListIndex = ListS�rala.ListIndex + 1
GoTo Bas
Son:
't�m dosyalar� kapat�r.
Close #1
'comboyu yeniler
ModelYenile
Exit Sub
Hata:
MsgBox "�stedi�iniz islem ger�eklestirilmedi."
End Sub
Private Sub Command11_Click()
ListS�rala.Clear
On Error GoTo Hata
'e�er combonun ilk maddesi se�ili (yani "<Bilinmeyen>") ise hata yapt�r.
If Combo3.ListIndex = 0 Then GoTo Hata
Soru = InputBox(Combo3.Text + " adl� t�r�n yeni ad�n� giriniz.", "T�r ad� de�istir")
If Len(Soru) = 0 Or Soru = String$(Len(Soru), " ") Then GoTo Hata
'bi ger�ek dosya bi de temp a��l�r.
Open App.Path + "\data\cmd5.nfo" For Input As #1
'd�ng� baslar, e�er ger�ek dosyadaki isim se�ili isim de�ilse temp e yaz
'aksi takdirde yazma ve di�erine ge�
ilk:
'dosyadaki t�m isimleri(de�isecek olan hari�) listeye ekler
'yeni ismi de ekler
If EOF(1) Then GoTo iki
Input #1, t�r
If Combo3.Text <> t�r Then
ListS�rala.AddItem t�r
GoTo ilk
Else
GoTo ilk
End If
iki:
Close #1
ListS�rala.AddItem UCase(Soru)
ListS�rala.ListIndex = 0
'sonra listedeki maddeleri dosyaya yazar
Open App.Path + "\data\cmd5.nfo" For Output As #1
Bas:
'yeni maddeyi ekler -listede se�ili olan�-
Write #1, ListS�rala.List(ListS�rala.ListIndex)
'e�er listenin sonuncusu se�iliyse bitir
If ListS�rala.ListIndex = ListS�rala.ListCount - 1 Then GoTo Son
'listenin bir alt�na ge�
ListS�rala.ListIndex = ListS�rala.ListIndex + 1
GoTo Bas
Son:
't�m dosyalar� kapat�r.
Close #1
'comboyu yeniler
T�rYenile
Exit Sub
Hata:
MsgBox "�stedi�iniz islem ger�eklestirilmedi."
End Sub
Private Sub Command12_Click()
ListS�rala.Clear
On Error GoTo Hata
'e�er combonun ilk maddesi se�ili (yani "<Bilinmeyen>") ise hata yapt�r.
If Combo4.ListIndex = 0 Then GoTo Hata
Soru = InputBox(Combo4.Text + " s�reli garantinin yeni s�resini giriniz.", "Garanti s�resi de�istir")
If Len(Soru) = 0 Or Soru = String$(Len(Soru), " ") Then GoTo Hata
'bi ger�ek dosya bi de temp a��l�r.
Open App.Path + "\data\cmd7.nfo" For Input As #1
'd�ng� baslar, e�er ger�ek dosyadaki isim se�ili isim de�ilse temp e yaz
'aksi takdirde yazma ve di�erine ge�
ilk:
'dosyadaki t�m isimleri(de�isecek olan hari�) listeye ekler
'yeni ismi de ekler
If EOF(1) Then GoTo iki
Input #1, gara
If Combo4.Text <> gara Then
ListS�rala.AddItem gara
GoTo ilk
Else
GoTo ilk
End If
iki:
Close #1
ListS�rala.AddItem UCase(Soru)
ListS�rala.ListIndex = 0
'sonra listedeki maddeleri dosyaya yazar
Open App.Path + "\data\cmd7.nfo" For Output As #1
Bas:
'yeni maddeyi ekler -listede se�ili olan�-
Write #1, ListS�rala.List(ListS�rala.ListIndex)
'e�er listenin sonuncusu se�iliyse bitir
If ListS�rala.ListIndex = ListS�rala.ListCount - 1 Then GoTo Son
'listenin bir alt�na ge�
ListS�rala.ListIndex = ListS�rala.ListIndex + 1
GoTo Bas
Son:
't�m dosyalar� kapat�r.
Close #1
'comboyu yeniler
GarantiYenile
Exit Sub
Hata:
MsgBox "�stedi�iniz islem ger�eklestirilmedi."
End Sub
Private Sub Form_Activate()
't�m combolar�n ilk halini "bilinmeyen" yapar, s�f�rdan baslat�r
Form1.Combo1.ListIndex = "0"
Form1.Combo2.ListIndex = "0"
Form1.Combo3.ListIndex = "0"
Form1.Combo4.ListIndex = "0"
Text1 = "": Text2 = "": Text3 = "": Text4 = "": Text5 = ""
ZamanBelirt
Cmmd1.SetFocus
End Sub
Private Sub Form_Load()
On Error Resume Next
Background.Left = 3750
Background.Width = Me.Width - 3750
Background.Top = 0
Background.Height = Me.Height
'comboyu yeniler
MarkaYenile
T�rYenile
ModelYenile
GarantiYenile
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub Timer1_Timer()
'ipucunun de�isme zaman� geldi...
Form1.LabelKontrol.Caption = "0"
'ipucunu belirtmek i�in fonksiyon �a��r�yoruz.
�pucuYaz
Timer1.Interval = 4000
End Sub
