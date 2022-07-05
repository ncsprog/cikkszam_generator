Attribute VB_Name = "Function_Cikkosztály_Csere"
Option Explicit

Sub Cikkosztály_Csere()
Dim JelöltOsztály As String, Oszlop As Integer, O1 As Integer, Ox As Integer, _
Sor As Integer, S1 As Integer, Sx As Integer

If AppCikkek.ComboBox10 <> "" Then
Exit Sub
Else
JelöltOsztály = AppCikkek.ComboBox11.Value
O1 = Munka2.Range("b2").Column
Ox = Munka2.Range("j2").Column
S1 = 2
Sx = 10
For Oszlop = 1 To Ox Step 1
    For Sor = S1 To Sx Step 1
        If Munka2.Cells(Sor, Oszlop).Value = JelöltOsztály Then
            Munka2.Cells(Sor, Oszlop).Value = AppCikkek.TextBox20.Value
        End If
    Next
Next
End If



End Sub
