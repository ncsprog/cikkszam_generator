Attribute VB_Name = "Function_IDgener�l�s"
Option Explicit

Sub IDgener�l�s()

Munka1.Select

Dim Idrw As Integer, Idnr As Integer

Sheets("adatok").Select
Columns("a:a").Select
Selection.End(xlDown).Select
Idrw = ActiveCell.Row + 1
Idnr = ActiveCell + 1
Range("a" & Idrw) = Idnr

End Sub
