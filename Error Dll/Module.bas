Attribute VB_Name = "Module"
Dim LogDLL As New Error.Cl_Err
Public Sub HataLogger()
    LogDLL.MsgErr App.Path, App.EXEName, Screen.ActiveForm.Name, Screen.ActiveControl.Name, _
    Err.Number, Err.Description, Err.Source, Err.HelpContext, Err.HelpFile, Err.LastDllError
    End
End Sub


