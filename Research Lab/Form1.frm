VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Research Lab"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6180
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   3420
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   960
      Top             =   4020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Color"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   4020
      Width           =   855
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Font"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4020
      Width           =   855
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   3420
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   4500
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox txtGiden 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   4500
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1826
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":000C
   End
   Begin RichTextLib.RichTextBox txtGelen 
      Height          =   3195
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5636
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0090
   End
   Begin RichTextLib.RichTextBox txtLoad 
      Height          =   3195
      Left            =   3120
      TabIndex        =   3
      Top             =   180
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5636
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0114
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdColor_Click()
    On Error GoTo CancelError
    If txtGiden.SelLength = 0 Then txtGiden.SelStart = 0: txtGiden.SelLength = Len(txtGiden.Text)
    CDialog.Color = txtGiden.SelColor 'yazýnýn özellikleri diyalog kutusuna aktarýlýr...
    CDialog.Flags = cdlCCFullOpen: CDialog.ShowColor 'font diyalog penceresi açýlýr.
    txtGiden.SelColor = CDialog.Color 'diyalog penceresinin deðiþkenlerinin yazýya aktarýrýr
CancelError:
End Sub
Private Sub cmdFont_Click()
    On Error GoTo CancelError
    If txtGiden.SelLength = 0 Then txtGiden.SelStart = 0: txtGiden.SelLength = Len(txtGiden.Text)
    'yazýnýn özellikleri diyalog kutusuna aktarýlýr...
    CDialog.FontBold = txtGiden.SelBold
    CDialog.Color = txtGiden.SelColor
    CDialog.FontItalic = txtGiden.SelItalic
    CDialog.FontName = txtGiden.SelFontName
    CDialog.FontSize = txtGiden.SelFontSize
    CDialog.FontStrikethru = txtGiden.SelStrikeThru
    CDialog.FontUnderline = txtGiden.SelUnderline
    'font diyalog penceresi açýlýr.
    CDialog.Flags = cdlCFScreenFonts Or cdlCFEffects: CDialog.ShowFont
    'diyalog penceresinin deðiþkenlerinin yazýya aktarýrýr
    txtGiden.SelBold = CDialog.FontBold
    txtGiden.SelColor = CDialog.Color
    txtGiden.SelItalic = CDialog.FontItalic
    txtGiden.SelFontName = CDialog.FontName
    txtGiden.SelFontSize = CDialog.FontSize
    txtGiden.SelStrikeThru = CDialog.FontStrikethru
    txtGiden.SelUnderline = CDialog.FontUnderline
CancelError:
End Sub
Private Sub cmdLoad_Click()
    txtLoad.LoadFile App.Path & "\chat.data"
End Sub
Private Sub cmdSave_Click()
    txtGiden.SaveFile App.Path & "\chat.data"
End Sub
Private Sub cmdSend_Click()
    txtGiden.Text = txtGiden.Text & vbCrLf & vbCrLf
    Clipboard.SetText txtGiden.Text: txtGelen.Text = txtGelen.Text & Clipboard.GetText
    Dim i As Integer: For i = 1 To Len(txtGiden.Text)
        txtGiden.SelStart = Len(txtGiden.Text) - i: txtGiden.SelLength = 1
        txtGelen.SelStart = Len(txtGelen.Text) - i: txtGelen.SelLength = 1
        txtGelen.SelBold = txtGiden.SelBold
        txtGelen.SelColor = txtGiden.SelColor
        txtGelen.SelItalic = txtGiden.SelItalic
        txtGelen.SelFontName = txtGiden.SelFontName
        txtGelen.SelFontSize = txtGiden.SelFontSize
        txtGelen.SelStrikeThru = txtGiden.SelStrikeThru
        txtGelen.SelUnderline = txtGiden.SelUnderline
    Next i
End Sub
Private Sub cmdCopy_Click()
    txtLoad.Text = txtLoad.Text & vbCrLf & vbCrLf
    Clipboard.SetText txtLoad.Text: txtGelen.Text = txtGelen.Text & Clipboard.GetText
    Dim i As Integer: For i = 1 To Len(txtLoad.Text)
        txtLoad.SelStart = Len(txtLoad.Text) - i: txtLoad.SelLength = 1
        txtGelen.SelStart = Len(txtGelen.Text) - i: txtGelen.SelLength = 1
        txtGelen.SelBold = txtLoad.SelBold
        txtGelen.SelColor = txtLoad.SelColor
        txtGelen.SelItalic = txtLoad.SelItalic
        txtGelen.SelFontName = txtLoad.SelFontName
        txtGelen.SelFontSize = txtLoad.SelFontSize
        txtGelen.SelStrikeThru = txtLoad.SelStrikeThru
        txtGelen.SelUnderline = txtLoad.SelUnderline
    Next i
End Sub

