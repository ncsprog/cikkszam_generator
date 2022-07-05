Attribute VB_Name = "Function_Rekord_Visszaadás"
Option Explicit

Sub Rekord_Visszaadás()

Munka1.Select

Dim Rw As Long, rngList As Range

Columns("a:a").Select
Selection.End(xlDown).Select
Rw = ActiveCell.Row

Set rngList = Munka1.Range("a3", "v" & Rw)
AppCikkek.ListBox1.List = rngList.Value
 

End Sub
