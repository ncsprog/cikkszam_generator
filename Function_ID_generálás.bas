Attribute VB_Name = "Function_ID_gener�l�s"
Option Explicit

Sub ID_gener�l�s()

Munka1.Select

Dim most As Date, ID_rw As Long
most = Now()

Munka1.Range("o1").Select
Columns("o:o").Select
Selection.End(xlDown).Select
ID_rw = ActiveCell.Row + 1
Range("o" & ID_rw).Value = most

End Sub
