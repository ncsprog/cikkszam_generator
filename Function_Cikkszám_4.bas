Attribute VB_Name = "Function_Cikksz�m_4"
Option Explicit

Sub Cikksz�m_4()

Munka1.Select

Dim El�tag As Integer, S1 As Integer, Sx As Long, MyRange As Range, Keresett As Integer, Ut�tag As Integer
S1 = Munka1.Range("p1").Row
Columns("p:p").Select
Selection.End(xlDown).Select
Sx = ActiveCell.Row
Keresett = Munka1.Range("p" & Sx).Value
Set MyRange = Munka1.Range("p" & S1, "p" & Sx)
El�tag = Application.WorksheetFunction.CountIf(MyRange, Keresett)

If El�tag > 999 Then
MsgBox "Ez a kateg�ria #999 rekordn�l betellt."
Munka1.Range("a" & Sx, "v" & Sx) = ""
Exit Sub
End If

If El�tag < 10 Then
Munka1.Range("q" & Sx).Value = Munka1.Range("p" & Sx).Value & "00" & El�tag
Else
End If

If El�tag > 9 Then
Munka1.Range("q" & Sx).Value = Munka1.Range("p" & Sx).Value & "0" & El�tag
Else
End If

If El�tag > 99 Then
Munka1.Range("q" & Sx).Value = Munka1.Range("p" & Sx).Value & El�tag
Else
End If

End Sub
