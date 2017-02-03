Attribute VB_Name = "sicil"
Sub Kayýtsýz()
MsgBox "   Kayýtlý Kiþi Yok     ", vbInformation, "3308 Sicil Takibi"
End Sub
Sub Ýptal()
Kayýtlar.Text1 = ""
Kayýtlar.Text2 = ""
Kayýtlar.Text3 = ""
Kayýtlar.Text4 = ""
Kayýtlar.Text1.Enabled = False
Kayýtlar.Text2.Enabled = False
Kayýtlar.Text3.Enabled = False
Kayýtlar.Text4.Enabled = False
Kayýtlar.Command1.Enabled = True
Kayýtlar.Command2.Enabled = False
Kayýtlar.Command10.Enabled = True
Kayýtlar.Command3.Enabled = True
Kayýtlar.Command9.Enabled = True
Kayýtlar.Command10.SetFocus
End Sub
Sub alanaeþitle()
Kayýtlar.Data1.Recordset.Fields("Ad Soyad") = Kayýtlar.Text1
Kayýtlar.Data1.Recordset.Fields("Doðum Tarihi") = Kayýtlar.Text2
Kayýtlar.Data1.Recordset.Fields("Sicil No") = Kayýtlar.Text3
Kayýtlar.Data1.Recordset.Fields("no") = Kayýtlar.Text4
Kayýtlar.Data1.Recordset.Update
End Sub
Sub ekranaeþitle()
Kayýtlar.Text1 = Kayýtlar.Data1.Recordset.Fields("Ad Soyad")
Kayýtlar.Text2 = Kayýtlar.Data1.Recordset.Fields("Doðum Tarihi")
Kayýtlar.Text3 = Kayýtlar.Data1.Recordset.Fields("Sicil No")
Kayýtlar.Text4 = Kayýtlar.Data1.Recordset.Fields("no")
End Sub
