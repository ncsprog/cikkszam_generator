Attribute VB_Name = "Function_Cikkszám_1"
Option Explicit

Sub Cikkszám_1()

Dim Fst As Integer, Snd As Integer, Trd As Integer, Foth As Integer, Fith As Integer, Sith As Integer, Seth As Integer, Eith As Integer, Nith As Integer
Fst = 1
Snd = 2
Trd = 3
Foth = 4
Fith = 5
Sith = 6
Seth = 7
Eith = 8
Nith = 9
' - elsõ tag
If AppCikkek.ComboBox2.Value = "" Then
Munka1.Range("x1").Value = 0
ElseIf AppCikkek.ComboBox2.Value = Munka2.Range("b4").Value Then
Munka1.Range("x1").Value = Fst
ElseIf AppCikkek.ComboBox2.Value = Munka2.Range("b5").Value Then
Munka1.Range("x1").Value = Snd
ElseIf AppCikkek.ComboBox2.Value = Munka2.Range("b6").Value Then
Munka1.Range("x1").Value = Trd
ElseIf AppCikkek.ComboBox2.Value = Munka2.Range("b7").Value Then
Munka1.Range("x1").Value = Foth
ElseIf AppCikkek.ComboBox2.Value = Munka2.Range("b8").Value Then
Munka1.Range("x1").Value = Fith
ElseIf AppCikkek.ComboBox2.Value = Munka2.Range("b9").Value Then
Munka1.Range("x1").Value = Sith
ElseIf AppCikkek.ComboBox2.Value = Munka2.Range("b10").Value Then
Munka1.Range("x1").Value = Seth
ElseIf AppCikkek.ComboBox2.Value = Munka2.Range("b11").Value Then
Munka1.Range("x1").Value = Eith
ElseIf AppCikkek.ComboBox2.Value = Munka2.Range("b12").Value Then
Munka1.Range("x1").Value = Nith
End If
End Sub
