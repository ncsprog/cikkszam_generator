Attribute VB_Name = "Function_Rekord_Visszaad�s"
Option Explicit

Sub Rekord_Visszaad�s()

Munka1.Select

Dim Rw As Long, rngList As Range

Columns("a:a").Select
Selection.End(xlDown).Select
Rw = ActiveCell.Row

Set rngList = Munka1.Range("a3", "v" & Rw)
AppCikkek.ListBox1.List = rngList.Value
 

End Sub
