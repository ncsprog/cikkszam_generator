Attribute VB_Name = "Function_Cikkt�rzs_Csere"
Option Explicit

Sub Cikkt�rzs_Csere()

Dim Jel�ltSor As String, Sor As Integer, S1 As Integer, Sn As Integer
If AppCikkek.TextBox20 = "" Then
Exit Sub
ElseIf AppCikkek.ComboBox11.Value <> "" Then
Exit Sub
ElseIf AppCikkek.ComboBox10.Value <> "" Then
Exit Sub
End If

Jel�ltSor = AppCikkek.ComboBox12.Value
S1 = Munka2.Range("a1").Cells.Row
Sn = S1 + 9

If Jel�ltSor = Null Then
MsgBox "Nincs kijel�lt sor."
Exit Sub
Else

For Sor = S1 To Sn Step 1
If Munka2.Range("a" & Sor).Value = Jel�ltSor Then
Munka2.Range("a" & Sor).Value = AppCikkek.TextBox20.Value
End If
Next
End If



End Sub
