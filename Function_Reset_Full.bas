Attribute VB_Name = "Function_Reset_Full"
Option Explicit

Sub Reset_Full()

Dim K�rd�s As Integer, Sx As Integer, K�rd�s2 As Integer, Rx As Long
Munka2.Select
Columns("dc:dc").Select
Selection.End(xlDown).Select
Sx = ActiveCell.Row

K�rd�s = MsgBox("Val�ban szeretn�d az aplik�ci� teljes tartalm�t alap�rt�kere �ll�tani?" & vbCrLf & vbCrLf & _
"Figyelem:" & vbCrLf & "Ezzel a folyamattal minden be�ll�tott megnevez�s alaphelyzetbe�ll, a mentett cikkadatok t�rl�dnek!", _
vbCritical + vbYesNo + vbDefaultButton2, "Alaphelyhzetbe �ll�t�s.")

If K�rd�s = vbYes Then
K�rd�s2 = MsgBox("A folyamat nem visszavonhat�! Biztosan folytatod?", vbExclamation + vbYesNo + vbDefaultButton2, "Program sz�vegeinek alaphelyzetbe �ll�t�sa.")
If K�rd�s2 = vbYes Then
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
