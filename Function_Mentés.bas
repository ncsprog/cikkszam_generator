Attribute VB_Name = "Function_Ment�s"
Option Explicit

Sub Ment�s()
' - k�sz

Munka1.Select

Dim Rwnr As Integer

Munka1.Select
Range("a1").Select
Selection.End(xlDown).Select
Rwnr = ActiveCell.Row

' - d�tum - B
Munka1.Range("b" & Rwnr).Value = Date
' - relevancia - C
Munka1.Range("c" & Rwnr).Value = AppCikkek.ComboBox1.Value
' - cikkt�rzs - D
Munka1.Range("d" & Rwnr).Value = AppCikkek.ComboBox2.Value
' - cikkoszt�ly - E
Munka1.Range("e" & Rwnr).Value = AppCikkek.ComboBox3.Value
' - cikkfaj - F
Munka1.Range("f" & Rwnr).Value = AppCikkek.ComboBox4.Value
' - st�tusz - G
Munka1.Range("g" & Rwnr).Value = AppCikkek.ComboBox5.Value
' - megnevez�s - H
Munka1.Range("h" & Rwnr).Value = AppCikkek.TextBox2.Value
' - megnevez�s - I
Munka1.Range("i" & Rwnr).Value = AppCikkek.TextBox3.Value
' - megnevez�s - J
Munka1.Range("j" & Rwnr).Value = AppCikkek.TextBox4.Value
' - megnevez�s - K
Munka1.Range("k" & Rwnr).Value = AppCikkek.TextBox5.Value
' - megnevez�s - L
Munka1.Range("l" & Rwnr).Value = AppCikkek.TextBox6.Value & ";" & AppCikkek.TextBox18.Value _
& ";" & AppCikkek.TextBox19.Value
' - megnevez�s - M
Munka1.Range("m" & Rwnr).Value = AppCikkek.TextBox7.Value
' - megnevez�s - N
Munka1.Range("n" & Rwnr).Value = AppCikkek.TextBox8.Value
' - megnevez�s - R
Munka1.Range("r" & Rwnr).Value = AppCikkek.TextBox11.Value
' - megnevez�s - S
Munka1.Range("s" & Rwnr).Value = AppCikkek.ComboBox6.Value
' - megnevez�s - T
Munka1.Range("t" & Rwnr).Value = AppCikkek.TextBox12.Value
' - megnevez�s - U
Munka1.Range("u" & Rwnr).Value = AppCikkek.TextBox13.Value
' - megnevez�s - V
Munka1.Range("v" & Rwnr).Value = AppCikkek.TextBox14.Value

AppCikkek.TextBox3 = ""
AppCikkek.TextBox4 = ""
AppCikkek.TextBox5 = ""
AppCikkek.TextBox6 = ""
AppCikkek.TextBox18 = ""
AppCikkek.TextBox19 = ""
AppCikkek.TextBox7 = ""
AppCikkek.TextBox8 = ""
AppCikkek.TextBox11 = ""
AppCikkek.TextBox12 = ""
AppCikkek.TextBox13 = ""
AppCikkek.TextBox14 = ""
AppCikkek.ComboBox6 = ""


End Sub
