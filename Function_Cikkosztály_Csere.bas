Attribute VB_Name = "Function_Cikkoszt�ly_Csere"
Option Explicit

Sub Cikkoszt�ly_Csere()
Dim Jel�ltOszt�ly As String, Oszlop As Integer, O1 As Integer, Ox As Integer, _
Sor As Integer, S1 As Integer, Sx As Integer

If AppCikkek.ComboBox10 <> "" Then
Exit Sub
Else
Jel�ltOszt�ly = AppCikkek.ComboBox11.Value
O1 = Munka2.Range("b2").Column
Ox = Munka2.Range("j2").Column
S1 = 2
Sx = 10
For Oszlop = 1 To Ox Step 1
    For Sor = S1 To Sx Step 1
        If Munka2.Cells(Sor, Oszlop).Value = Jel�ltOszt�ly Then
            Munka2.Cells(Sor, Oszlop).Value = AppCikkek.TextBox20.Value
        End If
    Next
Next
End If



End Sub
