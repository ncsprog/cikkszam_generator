Attribute VB_Name = "Function_Relevancia_Csere"
Option Explicit

Sub Relevancia_Csere()

Munka2.Select

Dim JelöltSor As String, Sor As Integer, S1 As Integer, Sn As Integer, Új As String
If AppCikkek.TextBox20.Value = "" Then
Exit Sub
End If

JelöltSor = AppCikkek.ComboBox13.Value
S1 = Munka2.Range("cu1").Cells.Row
Sn = S1 + 9

If JelöltSor = Null Then
MsgBox "Nincs kijelölt sor."
Exit Sub
Else

For Sor = S1 To Sn Step 1
If Munka2.Range("cu" & Sor).Value = JelöltSor Then
Munka2.Range("cu" & Sor).Value = AppCikkek.TextBox20.Value
End If
Next
End If


End Sub
