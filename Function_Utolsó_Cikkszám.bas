Attribute VB_Name = "Function_Utolsó_Cikkszám"
Option Explicit

Sub Utolsó_Cikkszám()

Dim Rw As Integer, Elõtag As Integer, Utótag As Integer

Munka1.Range("p1").Select
Columns("p:p").Select
Selection.End(xlDown).Select
Rw = ActiveCell.Row
Elõtag = Munka1.Range("p" & Rw)
Utótag = Munka1.Range("q" & Rw)

AppCikkek.TextBox10.Value = Elõtag & Utótag

End Sub
