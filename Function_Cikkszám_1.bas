Attribute VB_Name = "Function_Cikkszám_1"
Option Explicit

Sub Cikkszám_1()
' - kész, jó

Munka2.Select

Dim Sor As Integer, S1 As Integer, Sx As Integer, JelöltTörzs As String, TaláltSor As Integer
S1 = 2
Sx = 10
JelöltTörzs = AppCikkek.ComboBox2.Value

If AppCikkek.ComboBox2.Value <> "" Then
    For Sor = S1 To Sx Step 1
        If Munka2.Range("a" & Sor).Value = JelöltTörzs Then
            TaláltSor = Munka2.Range("a" & Sor).Row - 1
            End If
    Next
End If

Munka1.Select
Munka1.Range("w1").Value = TaláltSor
End Sub
