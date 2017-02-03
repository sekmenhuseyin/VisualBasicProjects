VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cl_Err"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub MsgErr(appPath As String, exeName As String, FormName As String, Procedure As String, Number As Integer, _
Description As String, Source As String, HelpContext As String, HelpFile As String, LastDllError As String)
    Dim iFileHandle As Integer: iFileHandle = FreeFile
    Open appPath & "\" & exeName & ".log" For Append As #iFileHandle
        Print #iFileHandle, "[ErrorDetails]"
        Print #iFileHandle, "Now=", Now
        Print #iFileHandle, "Source=", Source
        Print #iFileHandle, "FormName=", FormName
        Print #iFileHandle, "Procedure=", Procedure
        Print #iFileHandle, "Number=", Number
        Print #iFileHandle, "Description=", Description
        Print #iFileHandle, "HelpContext=", HelpContext
        Print #iFileHandle, "HelpFile=", HelpFile
        Print #iFileHandle, "LastDllError=", LastDllError
    Close #iFileHandle
    Open appPath & "\report.err" For Output As #iFileHandle
        Print #iFileHandle, "[ErrorDetails]"
        Print #iFileHandle, "Now=", Now
        Print #iFileHandle, "Source=", Source
        Print #iFileHandle, "FormName=", FormName
        Print #iFileHandle, "Procedure=", Procedure
        Print #iFileHandle, "Number=", Number
        Print #iFileHandle, "Description=", Description
        Print #iFileHandle, "HelpContext=", HelpContext
        Print #iFileHandle, "HelpFile=", HelpFile
        Print #iFileHandle, "LastDllError=", LastDllError
    Close #iFileHandle
    Shell appPath & "\ErrReport.exe", vbNormalFocus
End Sub
