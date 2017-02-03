VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Tazele"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "görünecek yazý"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1080
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mn1 
         Caption         =   "göster"
         Enabled         =   0   'False
      End
      Begin VB.Menu mn2 
         Caption         =   "gizle"
      End
      Begin VB.Menu mn3 
         Caption         =   "-"
      End
      Begin VB.Menu mn4 
         Caption         =   "kapat"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xState As Boolean
Private Sub Command1_Click()
    Show_SYStray Me, Text1.Text, "deðiþtir"
End Sub
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Form_Resize()
    If Me.WindowState = 1 And gizlendiMi = False Then
        mn2_Click
    ElseIf Me.WindowState = 0 Then
        xState = True
    ElseIf Me.WindowState = 2 Then
        xState = False
    End If
    Command1_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Show_SYStray Me, Text1.Text, "sil"
    End
End Sub
Private Sub Form_Load()
    gizlendiMi = False
    Show_SYStray Me, Text1.Text, "ekle"
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    menu.Visible = False
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu menu
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim msg As Long
    msg = X / Screen.TwipsPerPixelX
    ' Captura cada evento de botones del Raton
    Select Case msg
        Case WM_LBUTTONDBLCLK  ' Doble click Boton Izquierdo
        Case WM_LBUTTONDOWN  ' Boton Izquierdo pulsado
        Case WM_LBUTTONUP   ' Boton Izquierdo Soltado
            mn1_Click
        Case WM_RBUTTONDBLCLK ' Doble Click Boton Derecho
        Case WM_RBUTTONDOWN ' Boton derecho pulsado
        Case WM_RBUTTONUP  ' Boton derecho Arriba
            Me.PopupMenu Me.menu
    End Select
    DoEvents
End Sub
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub mn1_Click() 'Göster
    Me.Show
    If xState = True Then Me.WindowState = 0 Else Me.WindowState = 2
    gizlendiMi = False
    mn1.Enabled = False
    mn2.Enabled = True
End Sub
Private Sub mn2_Click() 'Sakla
    Me.Hide
    gizlendiMi = True
    mn1.Enabled = True
    mn2.Enabled = False
End Sub
Private Sub mn4_Click() 'Çýk
    Form_Unload (0)
End Sub

