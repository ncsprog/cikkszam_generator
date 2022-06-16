Attribute VB_Name = "Function_Mentés"
Option Explicit

Sub Mentés()

Dim RwNr As Integer

Munka1.Select
Range("a1").Select
Selection.End(xlDown).Select
RwNr = ActiveCell.Row + 1

' - dátum - B
Munka1.Range("b" & RwNr).Value = Date
' - relevancia - C
Munka1.Range("c" & RwNr).Value = AppCikkek.ComboBox1.Value
' - cikktörzs - D
Munka1.Range("d" & RwNr).Value = AppCikkek.ComboBox2.Value
' - cikkosztály - E
Munka1.Range("e" & RwNr).Value = AppCikkek.ComboBox3.Value
' - cikkfaj - F
Munka1.Range("f" & RwNr).Value = AppCikkek.ComboBox4.Value
' - státusz - G
Munka1.Range("g" & RwNr).Value = AppCikkek.ComboBox5.Value
' - megnevezés - H
Munka1.Range("h" & RwNr).Value = AppCikkek.TextBox2.Value
' - megnevezés - I
Munka1.Range("i" & RwNr).Value = AppCikkek.TextBox3.Value
' - megnevezés - J
Munka1.Range("j" & RwNr).Value = AppCikkek.TextBox4.Value
' - megnevezés - K
Munka1.Range("k" & RwNr).Value = AppCikkek.TextBox5.Value
' - megnevezés - L
Munka1.Range("l" & RwNr).Value = AppCikkek.TextBox6.Value
' - megnevezés - M
Munka1.Range("m" & RwNr).Value = AppCikkek.TextBox7.Value
' - megnevezés - N
Munka1.Range("n" & RwNr).Value = AppCikkek.TextBox8.Value

AppCikkek.TextBox3 = ""
AppCikkek.TextBox4 = ""
AppCikkek.TextBox5 = ""
AppCikkek.TextBox6 = ""
AppCikkek.TextBox7 = ""
AppCikkek.TextBox8 = ""


End Sub
