Attribute VB_Name = "Function_Cikkszám_4"
Option Explicit

Sub Cikkszám_4()

Munka1.Select

Dim Elõtag As Integer, S1 As Integer, Sx As Long, MyRange As Range, Keresett As Integer, Utótag As Integer
S1 = Munka1.Range("p1").Row
Columns("p:p").Select
Selection.End(xlDown).Select
Sx = ActiveCell.Row
Keresett = Munka1.Range("p" & Sx).Value
Set MyRange = Munka1.Range("p" & S1, "p" & Sx)
Elõtag = Application.WorksheetFunction.CountIf(MyRange, Keresett)

If Elõtag > 999 Then
MsgBox "Ez a kategória #999 rekordnál betellt."
Munka1.Range("a" & Sx, "v" & Sx) = ""
Exit Sub
End If

If Elõtag < 10 Then
Munka1.Range("q" & Sx).Value = Munka1.Range("p" & Sx).Value & "00" & Elõtag
Else
End If

If Elõtag > 9 Then
Munka1.Range("q" & Sx).Value = Munka1.Range("p" & Sx).Value & "0" & Elõtag
Else
End If

If Elõtag > 99 Then
Munka1.Range("q" & Sx).Value = Munka1.Range("p" & Sx).Value & Elõtag
Else
End If

End Sub
