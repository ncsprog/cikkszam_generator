Attribute VB_Name = "Function_Cikkszám_2"
Option Explicit

Sub Cikkszám_2()

Munka2.Select

Dim JelöltOsztály As String, Oszlop As Integer, O1 As Integer, Ox As Integer, _
Sor As Integer, S1 As Integer, Sx As Integer, TaláltSor As Integer


JelöltOsztály = AppCikkek.ComboBox3.Value
O1 = Munka2.Range("b2").Column
Ox = Munka2.Range("j2").Column
S1 = 2
Sx = 10
For Oszlop = 1 To Ox Step 1
    For Sor = S1 To Sx Step 1
        If Munka2.Cells(Sor, Oszlop).Value = JelöltOsztály Then
            TaláltSor = Munka2.Cells(Sor, Oszlop).Row - 1
            Munka1.Range("y1").Value = TaláltSor
        End If
    Next
Next


End Sub
