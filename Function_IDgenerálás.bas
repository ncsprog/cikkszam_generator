Attribute VB_Name = "Function_IDgenerálás"
Option Explicit

Sub IDgenerálás()

Munka1.Select

Dim Idrw As Integer, Idnr As Integer

Sheets("adatok").Select
Columns("a:a").Select
Selection.End(xlDown).Select
Idrw = ActiveCell.Row + 1
Idnr = ActiveCell + 1
Range("a" & Idrw) = Idnr

End Sub
