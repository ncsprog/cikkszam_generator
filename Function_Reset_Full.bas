Attribute VB_Name = "Function_Reset_Full"
Option Explicit

Sub Reset_Full()

Dim Kérdés As Integer, Sx As Integer, Kérdés2 As Integer, Rx As Long
Munka2.Select
Columns("dc:dc").Select
Selection.End(xlDown).Select
Sx = ActiveCell.Row

Kérdés = MsgBox("Valóban szeretnéd az aplikáció teljes tartalmát alapértékere állítani?" & vbCrLf & vbCrLf & _
"Figyelem:" & vbCrLf & "Ezzel a folyamattal minden beállított megnevezés alaphelyzetbeáll, a mentett cikkadatok törlõdnek!", _
vbCritical + vbYesNo + vbDefaultButton2, "Alaphelyhzetbe állítás.")

If Kérdés = vbYes Then
Kérdés2 = MsgBox("A folyamat nem visszavonható! Biztosan folytatod?", vbExclamation + vbYesNo + vbDefaultButton2, "Program szövegeinek alaphelyzetbe állítása.")
If Kérdés2 = vbYes Then
Munka2.Range("df2", "df" & Sx).Copy
Munka2.Range("dc2").PasteSpecial xlPasteValues
Munka2.Range("a30:cq38").Copy
Munka2.Range("a2").PasteSpecial xlPasteValues
Munka2.Range("cu30:cw38").Copy
Munka2.Range("cu2").PasteSpecial xlPasteValues



Munka1.Select

Columns("a:a").Select
Selection.End(xlDown).Select
Rx = ActiveCell.Row
Munka1.Range("a2", "w" & Rx) = ""
Munka1.Range("a2:w2") = "0"
Else
Exit Sub
End If
End If
End Sub
