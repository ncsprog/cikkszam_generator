Attribute VB_Name = "Function_Cikkfaj_Csere"
Option Explicit

Sub Cikkfaj_Csere()
Dim Jel�ltFaj As String, Oszlop As Integer, O1 As Integer, Ox As Integer, _
Sor As Integer, S1 As Integer, Sx As Integer, Jel�ltOszt�ly As String

Jel�ltFaj = AppCikkek.ComboBox10.Value
O1 = Munka2.Range("k2").Column
Ox = Munka2.Range("cm2").Column
S1 = 2
Sx = 10
For Oszlop = 1 To Ox Step 1
    For Sor = S1 To Sx Step 1
        If Munka2.Cells(Sor, Oszlop).Value = Jel�ltFaj Then
            Munka2.Cells(Sor, Oszlop).Value = AppCikkek.TextBox20.Value
        End If
    Next
Next


End Sub
