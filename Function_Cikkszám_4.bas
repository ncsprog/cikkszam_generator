Attribute VB_Name = "Function_Cikksz�m_4"
Option Explicit

Sub Cikksz�m_4()

Dim Rw As Integer, El�tag As Integer, Db As Integer, Ut�tag As Integer, Ut�H As Integer

El�tag = Munka1.Range("x1").Value & Munka1.Range("y1").Value & Munka1.Range("z1").Value

Munka1.Range("p1").Select
Columns("p:p").Select
Selection.End(xlDown).Select
Rw = ActiveCell.Row + 1

Db = Application.WorksheetFunction.CountIf(Range("p3", "p" & Rw), El�tag)
Ut�tag = Db + 1
Ut�H = Len(Db)
Munka1.Range("p" & Rw).Value = El�tag

If Ut�tag > 950 Then
MsgBox "Hamarosan el�ri a maximum darabsz�mot ez a Cikkfaj!)"
End If
If Ut�tag = 1000 Then
MsgBox "Elfogyott a cikkt�rhely, keress m�sik cikkoszt�lyt!"
Exit Sub
End If
If Ut�H = 1 Then
Munka1.Range("q" & Rw).Value = "Kar" & El�tag & "00" & Ut�tag
ElseIf Ut�H = 2 Then
Munka1.Range("q" & Rw).Value = "Kar" & El�tag & "0" & Ut�tag
ElseIf Ut�H = 3 Then
Munka1.Range("q" & Rw).Value = "Kar" & El�tag & Ut�tag
End If

End Sub
