Attribute VB_Name = "Function_Reset"
Option Explicit

Sub Reset()



Dim K�rd�s As Integer, Sx
Munka2.Select
Columns("dc:dc").Select
Selection.End(xlDown).Select
Sx = ActiveCell.Row

K�rd�s = MsgBox("Val�ban szeretn�d az �sszes feliratot alaphelyzetbe �ll�tani?", vbExclamation + vbYesNo + vbDefaultButton2, "Program sz�vegeinek alaphelyzetbe �ll�t�sa.")

If K�rd�s = vbYes Then
Munka2.Range("df2", "df" & Sx).Copy
Munka2.Range("dc2").PasteSpecial xlPasteValues
Else
Exit Sub
End If

End Sub
