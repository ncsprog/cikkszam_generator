Attribute VB_Name = "Function_El�tag_Csere"
Option Explicit

Sub El�tag_Csere()

Munka2.Select

Dim Jel�ltSor As String, Sor As Integer, S1 As Integer, Sn As Integer
If AppCikkek.TextBox20 = "" Then
Exit Sub
End If

Jel�ltSor = AppCikkek.ComboBox14.Value
S1 = Munka2.Range("co1").Cells.Row
Sn = S1 + 9

If Jel�ltSor = Null Then
MsgBox "Nincs kijel�lt sor."
Exit Sub
Else

For Sor = S1 To Sn Step 1
If Munka2.Range("co" & Sor).Value = Jel�ltSor Then
Munka2.Range("co" & Sor).Value = AppCikkek.TextBox20.Value
End If
Next
End If


End Sub
