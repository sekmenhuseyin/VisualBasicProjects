Attribute VB_Name = "Module"
Option Explicit
Global MyName As String '[Að]
Sub Main()
    MyName = Environ("COMPUTERNAME")
    frmMain.Show
End Sub

