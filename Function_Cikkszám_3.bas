Attribute VB_Name = "Function_Cikkszám_3"
Option Explicit

Sub Cikkszám_3()
' - kész, jó

Munka2.Select

Dim JelöltFaj As String, Oszlop As Integer, O1 As Integer, Ox As Integer, _
Sor As Integer, S1 As Integer, Sx As Integer, TaláltSor As Integer, ElõtagSor As Long


JelöltFaj = AppCikkek.ComboBox4.Value
O1 = Munka2.Range("k2").Column
Ox = Munka2.Range("cm2").Column
S1 = 2
Sx = 10
For Oszlop = 1 To Ox Step 1
    For Sor = S1 To Sx Step 1
        If Munka2.Cells(Sor, Oszlop).Value = JelöltFaj Then
            TaláltSor = Munka2.Cells(Sor, Oszlop).Row - 1
            Munka1.Range("z1").Value = TaláltSor
        End If
    Next
Next

Munka1.Select

Columns("o:o").Select
Selection.End(xlDown).Select
ElõtagSor = ActiveCell.Row
Munka1.Range("p" & ElõtagSor).Value = Munka1.Range("x1").Value & Munka1.Range("y1").Value & Munka1.Range("z1").Value


End Sub
