Attribute VB_Name = "Function_Rekord_Visszaad�s"
Option Explicit

Sub Rekord_Visszaad�s()

Dim Rw As Integer, rngList As Range

Munka1.Range("a1").Select
Columns("a:a").Select
Selection.End(xlDown).Select
Rw = ActiveCell.Row

Set rngList = Munka1.Range("a1", "v" & Rw)
AppCikkek.ListBox1.List = rngList.Value
 

End Sub
