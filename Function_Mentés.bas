Attribute VB_Name = "Function_Ment�s"
Option Explicit

Sub Ment�s()

Dim RwNr As Integer

Munka1.Select
Range("a1").Select
Selection.End(xlDown).Select
RwNr = ActiveCell.Row + 1

' - d�tum - B
Munka1.Range("b" & RwNr).Value = Date
' - relevancia - C
Munka1.Range("c" & RwNr).Value = AppCikkek.ComboBox1.Value
' - cikkt�rzs - D
Munka1.Range("d" & RwNr).Value = AppCikkek.ComboBox2.Value
' - cikkoszt�ly - E
Munka1.Range("e" & RwNr).Value = AppCikkek.ComboBox3.Value
' - cikkfaj - F
Munka1.Range("f" & RwNr).Value = AppCikkek.ComboBox4.Value
' - st�tusz - G
Munka1.Range("g" & RwNr).Value = AppCikkek.ComboBox5.Value
' - megnevez�s - H
Munka1.Range("h" & RwNr).Value = AppCikkek.TextBox2.Value
' - megnevez�s - I
Munka1.Range("i" & RwNr).Value = AppCikkek.TextBox3.Value
' - megnevez�s - J
Munka1.Range("j" & RwNr).Value = AppCikkek.TextBox4.Value
' - megnevez�s - K
Munka1.Range("k" & RwNr).Value = AppCikkek.TextBox5.Value
' - megnevez�s - L
Munka1.Range("l" & RwNr).Value = AppCikkek.TextBox6.Value
' - megnevez�s - M
Munka1.Range("m" & RwNr).Value = AppCikkek.TextBox7.Value
' - megnevez�s - N
Munka1.Range("n" & RwNr).Value = AppCikkek.TextBox8.Value

AppCikkek.TextBox3 = ""
AppCikkek.TextBox4 = ""
AppCikkek.TextBox5 = ""
AppCikkek.TextBox6 = ""
AppCikkek.TextBox7 = ""
AppCikkek.TextBox8 = ""


End Sub
