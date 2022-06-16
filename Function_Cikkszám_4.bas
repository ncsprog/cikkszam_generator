Attribute VB_Name = "Function_Cikkszám_4"
Option Explicit

Sub Cikkszám_4()

Dim Rw As Integer, Elõtag As Integer, Db As Integer, Utótag As Integer, UtóH As Integer

Elõtag = Munka1.Range("x1").Value & Munka1.Range("y1").Value & Munka1.Range("z1").Value

Munka1.Range("p1").Select
Columns("p:p").Select
Selection.End(xlDown).Select
Rw = ActiveCell.Row + 1

Db = Application.WorksheetFunction.CountIf(Range("p3", "p" & Rw), Elõtag)
Utótag = Db + 1
UtóH = Len(Db)
Munka1.Range("p" & Rw).Value = Elõtag

If Utótag > 950 Then
MsgBox "Hamarosan eléri a maximum darabszámot ez a Cikkfaj!)"
End If
If Utótag = 1000 Then
MsgBox "Elfogyott a cikktárhely, keress másik cikkosztályt!"
Exit Sub
End If
If UtóH = 1 Then
Munka1.Range("q" & Rw).Value = "Kar" & Elõtag & "00" & Utótag
ElseIf UtóH = 2 Then
Munka1.Range("q" & Rw).Value = "Kar" & Elõtag & "0" & Utótag
ElseIf UtóH = 3 Then
Munka1.Range("q" & Rw).Value = "Kar" & Elõtag & Utótag
End If

End Sub
