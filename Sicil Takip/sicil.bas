Attribute VB_Name = "sicil"
Sub Kay�ts�z()
MsgBox "   Kay�tl� Ki�i Yok     ", vbInformation, "3308 Sicil Takibi"
End Sub
Sub �ptal()
Kay�tlar.Text1 = ""
Kay�tlar.Text2 = ""
Kay�tlar.Text3 = ""
Kay�tlar.Text4 = ""
Kay�tlar.Text1.Enabled = False
Kay�tlar.Text2.Enabled = False
Kay�tlar.Text3.Enabled = False
Kay�tlar.Text4.Enabled = False
Kay�tlar.Command1.Enabled = True
Kay�tlar.Command2.Enabled = False
Kay�tlar.Command10.Enabled = True
Kay�tlar.Command3.Enabled = True
Kay�tlar.Command9.Enabled = True
Kay�tlar.Command10.SetFocus
End Sub
Sub alanae�itle()
Kay�tlar.Data1.Recordset.Fields("Ad Soyad") = Kay�tlar.Text1
Kay�tlar.Data1.Recordset.Fields("Do�um Tarihi") = Kay�tlar.Text2
Kay�tlar.Data1.Recordset.Fields("Sicil No") = Kay�tlar.Text3
Kay�tlar.Data1.Recordset.Fields("no") = Kay�tlar.Text4
Kay�tlar.Data1.Recordset.Update
End Sub
Sub ekranae�itle()
Kay�tlar.Text1 = Kay�tlar.Data1.Recordset.Fields("Ad Soyad")
Kay�tlar.Text2 = Kay�tlar.Data1.Recordset.Fields("Do�um Tarihi")
Kay�tlar.Text3 = Kay�tlar.Data1.Recordset.Fields("Sicil No")
Kay�tlar.Text4 = Kay�tlar.Data1.Recordset.Fields("no")
End Sub
