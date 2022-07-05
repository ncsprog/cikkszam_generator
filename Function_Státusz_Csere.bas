Attribute VB_Name = "Function_Státusz_Csere"
Option Explicit

Sub Státusz_Csere()

Dim JelöltSor As String, Sor As Integer, S1 As Integer, Sn As Integer
If AppCikkek.TextBox20 = "" Then
Exit Sub
End If

JelöltSor = AppCikkek.ComboBox9.Value
S1 = Munka2.Range("cw1").Cells.Row
Sn = S1 + 9

If JelöltSor = Null Then
MsgBox "Nincs kijelölt sor."
Exit Sub
Else

For Sor = S1 To Sn Step 1
If Munka2.Range("cw" & Sor).Value = JelöltSor Then
Munka2.Range("cw" & Sor).Value = AppCikkek.TextBox20.Value
End If
Next
End If

End Sub
