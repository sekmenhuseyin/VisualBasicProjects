VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{32BFFBBF-2161-43EE-B99C-F043EF1F948F}#1.0#0"; "SENXPCTL.ocx"
Begin VB.Form CopyFile 
   BackColor       =   &H0096E06D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yedekle"
   ClientHeight    =   4005
   ClientLeft      =   1695
   ClientTop       =   1515
   ClientWidth     =   5340
   ClipControls    =   0   'False
   Icon            =   "frmCopy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4005
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmr_Geciktirici 
      Interval        =   100
      Left            =   2100
      Top             =   4305
   End
   Begin VB.Frame Frame 
      BackColor       =   &H0096E06D&
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   135
      Width           =   5100
      Begin ComctlLib.ProgressBar ProgressBar 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2310
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.TextBox Filepath 
         Height          =   600
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox Destinationpath 
         Height          =   600
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1500
         Width           =   3255
      End
      Begin OsenXPCntrl.OsenXPButton Browsefile 
         Height          =   840
         Left            =   3465
         TabIndex        =   0
         Top             =   240
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1482
         BTYPE           =   3
         TX              =   "Gözat"
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
         MICON           =   "frmCopy.frx":0442
         PICN            =   "frmCopy.frx":045E
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.OsenXPButton Browsedestination 
         Height          =   840
         Left            =   3465
         TabIndex        =   1
         Top             =   1260
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1482
         BTYPE           =   3
         TX              =   "Gözat"
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
         MICON           =   "frmCopy.frx":08B0
         PICN            =   "frmCopy.frx":08CC
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Percentlabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Tamamlanma Yüzdesi"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1545
         Width           =   3855
      End
      Begin VB.Label Filelabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Kaynak"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Destinationlabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Hedef"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1260
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   0
      Top             =   4410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   6148
   End
   Begin OsenXPCntrl.OsenXPButton Cancel 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   3135
      Width           =   1995
      _ExtentX        =   3519
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
      MICON           =   "frmCopy.frx":0D1E
      PICN            =   "frmCopy.frx":0D3A
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   2
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton Copy 
      Default         =   -1  'True
      Height          =   735
      Left            =   3225
      TabIndex        =   2
      Top             =   3135
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Kopyala"
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
      MICON           =   "frmCopy.frx":118C
      PICN            =   "frmCopy.frx":11A8
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
Attribute VB_Name = "CopyFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Type BROWSEINFO: hOwner As Long: pidlRoot As Long: pszDisplayName As String: lpszTitle As String: ulFlags As Long: lpfn As Long: lParam As Long: iImage As Long: End Type
Private Type SHITEMID: cb As Long: abID As Byte: End Type
Private Type ITEMIDLIST: mkid As SHITEMID: End Type
Private Const NOERROR = 0: Private Const BIF_RETURNONLYFSDIRS = &H1: Private Const BIF_DONTGOBELOWDOMAIN = &H2: Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8: Private Const BIF_BROWSEFORCOMPUTER = &H1000: Private Const BIF_BROWSEFORPRINTER = &H2000
Private Sub Browsedestination_Click()
    On Local Error Resume Next
    Dim bi As BROWSEINFO 'declare the needed variables
    Dim rtn&, pidl&, path$, pos%
    Dim t As Long   'temp
    Dim SpecIn, SpecOut As String
    bi.hOwner = Me.hwnd 'centres the dialog on the screen
    bi.lpszTitle = "Hedef Klasör" 'set the title text
    bi.ulFlags = BIF_RETURNONLYFSDIRS 'the type of folder(s) to return
    pidl& = SHBrowseForFolder(bi) 'show the dialog box
    path = Space(512) 'sets the maximum characters
    t = SHGetPathFromIDList(ByVal pidl&, ByVal path) 'gets the selected path
    pos% = InStr(path$, Chr$(0)) 'extracts the path from the string
    SpecIn = Left(path$, pos - 1) 'sets the extracted path to SpecIn
    If Right$(SpecIn, 1) = "\" Then SpecOut = SpecIn Else SpecOut = SpecIn + "\"
    Destinationpath.Text = SpecOut + ExtractName(Filepath.Text) 'merges both the destination path and the source filename into one string
End Sub
Private Sub Browsefile_Click()
    Dialog.FileName = App.path + "\include\data\Data.mdb"
    Dialog.Filter = "Veritabaný Dosyalarý|*.mdb"
    Dialog.DialogTitle = "Kaynak Veritabaný" 'set the dialog title
    Dialog.ShowOpen 'show the dialog box
    Filepath.Text = Dialog.FileName 'set the target text box to the file chosen
End Sub
Private Sub Cancel_Click()
    Unload Me
End Sub
Private Sub Copy_Click()
    Write #7, Time$, Me.Name, "Copy_Click", "Start" 'logging
    On Error GoTo hataLog
    If Filepath.Text = "" Or Dir(Filepath.Text) = "" Then 'make sure that a target file is specified
        MsgBox "Hedef dosyanýn yolunu va adýný belitmelisiniz.", vbCritical 'if not then display a message
        Exit Sub                                                                       'and exit the procedure
    End If
    If Destinationpath.Text = "" Then 'make sure that a destination path is specified
        MsgBox "Kaynak dosyanýn yolunu ve adýný belirtmelisiniz.", vbCritical 'if not then display a message
        Exit Sub                                                                          'and exit the procedure
    End If
    Filepath.Locked = True: Destinationpath.Locked = True: Browsefile.Enabled = False: Browsedestination.Enabled = False: Copy.Enabled = False: Cancel.Enabled = False
    'if all is OK then copy the file
    If Me.Caption = "Veritabanýný Diskete Yedekle" And Dir("a:\" & CStr(Date) & "\") = "" Then MkDir ("a:\" & CStr(Date) & "\")
    CopyFile Filepath.Text, Destinationpath.Text
    ProgressBar.Value = 0 'returns the progress bar to zero
    If Me.Caption = "Veritabanýný Yedekten Geri Al" Then
        Filepath.Locked = False
        Browsefile.Enabled = True
    Else
        Destinationpath.Locked = False
        Browsedestination.Enabled = True
    End If
    Copy.Enabled = True: Cancel.Enabled = True: Cancel.SetFocus
    Write #7, Time$, Me.Name, "Copy_Click", "From:" & Filepath.Text & " & To:" & Destinationpath.Text 'logging
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Sub
hataLog:
    Select Case Err.Number
        Case 52
            MsgBox "Lütfen bir disket yerleþtirin.", vbExclamation + vbRetryCancel, "Disket Yok"
        Case Else
            MsgBox "Kopyalama iþlemi baþarýsýzlýkla sonuçlandý.", vbExclamation
    End Select
    Write #7, Time$, Me.Name, "Copy_Click", "UnSuccessful" 'logging
End Sub
Private Sub Form_Load()
    Write #7, Time$, Me.Name, "Form_Load", "Start" 'logging
    Dim Control As Control: Me.Move (frmMain.ScaleWidth - Width) / 2, (frmMain.ScaleHeight - Height) / 2
    Me.BackColor = rnk_frm_arka: Frame.BackColor = rnk_frm_arka
    Filepath.BackColor = rnk_yazý_arka: Filepath.ForeColor = rnk_yazý_ön: Destinationpath.BackColor = rnk_yazý_arka: Destinationpath.ForeColor = rnk_yazý_ön
    For Each Control In Me
        If TypeOf Control Is OsenXPButton Then Control.BackColor = rnk_btn_arka: Control.BackOver = rnk_btn_arka: Control.ForeColor = rnk_btn_ön: Control.ForeOver = rnk_btn_ön
        If TypeOf Control Is Label Then Control.ForeColor = rnk_yazý_ön
    Next Control
    Write #7, Time$, Me.Name, "Form_Load", "Successful" 'logging
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Write #7, Time$, Me.Name, "Form_Unload", "Successful" 'logging
    Ayarlar.Enabled = True: Ayarlar.Command1.SetFocus
End Sub
Private Sub tmr_Geciktirici_Timer()
    Write #7, Time$, Me.Name, "tmr_Geciktirici_Timer", "Start" 'logging
    Select Case Me.Caption
        Case "1"    'Veritabanýný Yedekle
            Me.Caption = "Veritabanýný Yedekle"
            Filepath.Text = App.path + "\include\data\Data.mdb": Filepath.Locked = True: Browsefile.Enabled = False
        Case "2"    'Veritabanýný Yedekten Geri Al
            Me.Caption = "Veritabanýný Yedekten Geri Al"
            Destinationpath.Text = App.path + "\include\data\Data.mdb": Destinationpath.Locked = True: Browsedestination.Enabled = False
        Case "3"    'Veritabanýný Diskete Yedekle
            Me.Caption = "Veritabanýný Diskete Yedekle"
            Filepath.Text = App.path + "\include\data\Data.mdb": Filepath.Locked = True: Browsefile.Enabled = False
            Destinationpath.Text = "a:\" & CStr(Date) & "\include\data\Data.mdb": Destinationpath.Locked = True: Browsedestination.Enabled = False
    End Select
    SendKeys "{enter}"
    tmr_Geciktirici.Enabled = False
    Write #7, Time$, Me.Name, "tmr_Geciktirici_Timer", "End" 'logging
End Sub

'This project was downloaded from
'
'http://www.brianharper.demon.co.uk/
'
'Please use this project and all of its source code however you want.
'
'UNZIPPING
'To unzip the project files you will need a 32Bit unzipper program that
'can handle long file names. If you have a latest copy of Winzip installed
'on your system then you may use that. If you however dont have a copy,
'then visit my web site, go into the files section and from there you can
'click on the Winzip link to goto their site and download a copy of the
'program. By doing this you will now beable to unzip the project files
'retaining their proper long file names.
'Once upzipped, load up your copy of Visual Basic and goto
'File/Open Project. Locate the project files to where ever you unzipped
'them, then click Open. The project files will be loaded and are now ready
'for use.
'
'THE PROJECT
'I created this project in order to try and spice up a menu system I was
'once working on. I needed to copy files betweem disks and needed some
'indication of how long it would take and how it was doing. Using a percent
'bar in the project would have been ideal. Percent bars are now used as a
'common method of indicating how a procedure is doing. They might not be
'100% accurate but they are the next best thing. After hours of research
'and many hours of debugging, I finally came up with an easy to use
'executable using a percent bar while copying a file, which was ideally
'suited to what I needed.
'
'NOTES
'I have only provided the necessary project files with the zip. This keeps
'the size of the zip files down to a minimum and enables me to upload more
'prjects files to my site.
'
'I hope you find the project usful in what ever you are programming. I
'have tried to write out a small explanation of what each line of code
'does in the project, although most of it is pretty simple to understand.
'
'If you find any bugs in the code then please dont hesitate to Email me and
'I will get back to you as soon as possible. If you however need help on a
'different matter concerning Visual Basic then please please Email me as
'I like to here from people and here what they are programming.
'
'My Email address is:
'Brian@brianharper.demon.co.uk
'
'My web site is:
'http://www.brianharper.demon.co.uk/
'
'Please visit my web site and find many other useful projects like this.
'

'This code is used to copy the file provided in the Source text box. The
'file is calculated and then copied to the destination path while advancing
'the progress bar at the same time.
Function CopyFile(Src As String, dst As String) As Single
    Write #7, Time$, Me.Name, "CopyFile", "Start" 'logging
    On Local Error GoTo hataLog
    Static Buf$
    Dim BTest!, FSize! 'declare the needed variables
    Dim Chunk%, F1%, F2%
    Const BUFSIZE = 1024 'set the buffer size
    If Len(Dir(dst)) Then 'check to see if the destination file already exists
        If MsgBox(dst + Chr(10) + Chr(10) + "Dosya zaten mevcut. Üzerine yazmak istiyor musunuz?", vbYesNo + vbQuestion) = vbNo Then Exit Function Else Kill dst
    End If
    F1 = FreeFile 'returns file number available
    Open Src For Binary As F1 'open the source file
    F2 = FreeFile 'returns file number available
    Open dst For Binary As F2 'open the destination file
    FSize = LOF(F1)
    BTest = FSize - LOF(F2)
    Do
        If BTest < BUFSIZE Then
            Chunk = BTest
        Else
            Chunk = BUFSIZE
        End If
        Buf = String(Chunk, " ")
        Get F1, , Buf
        Put F2, , Buf
        BTest = FSize - LOF(F2)
        If (100 - Int(100 * BTest / FSize)) > 100 Then ProgressBar.Value = 100 Else ProgressBar.Value = (100 - Int(100 * BTest / FSize)) 'advance the progress bar as the file is copied
    Loop Until BTest = 0
    Close F1 'closes the source file
    Close F2 'closes the destination file
    CopyFile = FSize
    Write #7, Time$, Me.Name, "CopyFile", "Successful" 'logging
    MsgBox "Kopyalama baþarýyla gerçekleþmiþtir.", vbInformation, "Yedekleme"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Function
hataLog:
    Close F1: Close F2: ProgressBar.Value = 0
    Select Case Err.Number
        Case 52, 57, 31036: MsgBox "Lütfen bir disket yerleþtirin.", vbExclamation
        Case 61: MsgBox "Kayda devam edecek boþ alan kalmadý.", vbExclamation
        Case 71: MsgBox "Disk hazýr deðil.", vbExclamation
        Case 75, 70: MsgBox "Disk salt okunur.", vbExclamation
        Case 76: MsgBox "Yol bulunamadý.", vbExclamation
        Case 320: MsgBox "Hatalý bir dosya adý seçtiniz.", vbExclamation
        Case Else:  MsgBox "Kopyalama iþlemi baþarýsýzlýkla sonuçlandý.", vbExclamation
    End Select
    Write #7, Time$, Me.Name, "CopyFile", "Unsuccessful" 'logging
End Function
'This code is used to extract the filename provided by the user from the
'Source text box. The filename is extracted and passed to the string
'SpecOut. Once the filename is extraced from the text box, it is then added
'to the destination path provided by the user.
Private Function ExtractName(SpecIn As String) As String
    Dim i As Integer 'declare the needed variables
    Dim SpecOut As String
    On Local Error Resume Next 'ignore any errors
    For i = Len(SpecIn) To 1 Step -1 ' assume what follows the last backslash is the file to be extracted
        If Mid(SpecIn, i, 1) = "\" Then
            SpecOut = Mid(SpecIn, i + 1) 'extract the filename from the path provided
            Exit For
        End If
    Next i
    ExtractName = SpecOut 'returns the extracted filename from the path
End Function

