Attribute VB_Name = "Function_Utols�_Cikksz�m"
Option Explicit

Sub Utols�_Cikksz�m()

Dim Rw As Integer, El�tag As Integer, Ut�tag As Integer

Munka1.Range("p1").Select
Columns("p:p").Select
Selection.End(xlDown).Select
Rw = ActiveCell.Row
El�tag = Munka1.Range("p" & Rw)
Ut�tag = Munka1.Range("q" & Rw)

AppCikkek.TextBox10.Value = El�tag & Ut�tag

End Sub
