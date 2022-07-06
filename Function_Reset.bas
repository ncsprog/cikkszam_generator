Attribute VB_Name = "Function_Reset"
Option Explicit

Sub Reset()



Dim Kérdés As Integer, Sx
Munka2.Select
Columns("dc:dc").Select
Selection.End(xlDown).Select
Sx = ActiveCell.Row

Kérdés = MsgBox("Valóban szeretnéd az összes feliratot alaphelyzetbe állítani?", vbExclamation + vbYesNo + vbDefaultButton2, "Program szövegeinek alaphelyzetbe állítása.")

If Kérdés = vbYes Then
Munka2.Range("df2", "df" & Sx).Copy
Munka2.Range("dc2").PasteSpecial xlPasteValues
Else
Exit Sub
End If

End Sub
