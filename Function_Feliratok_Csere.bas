Attribute VB_Name = "Function_Feliratok_Csere"
Option Explicit

Sub Feliratok_Csere()

Munka2.Select

Dim JelöltSor As String, Sor As Integer, S1 As Integer, Sn As Integer
If AppCikkek.TextBox20 = "" Then
Exit Sub
End If

JelöltSor = AppCikkek.ComboBox16.Value
S1 = Munka2.Range("dc1").Cells.Row
Columns("dc:dc").Select
Selection.End(xlDown).Select
Sn = ActiveCell.Row

If JelöltSor = Null Then
MsgBox "Nincs kijelölt sor."
Exit Sub
Else

For Sor = S1 To Sn Step 1
If Munka2.Range("dc" & Sor).Value = JelöltSor Then
Munka2.Range("dc" & Sor).Value = AppCikkek.TextBox20.Value
End If
Next
End If

End Sub
