Attribute VB_Name = "Function_Cikksz�m_1"
Option Explicit

Sub Cikksz�m_1()
' - k�sz, j�

Munka2.Select

Dim Sor As Integer, S1 As Integer, Sx As Integer, Jel�ltT�rzs As String, Tal�ltSor As Integer
S1 = 2
Sx = 10
Jel�ltT�rzs = AppCikkek.ComboBox2.Value

If AppCikkek.ComboBox2.Value <> "" Then
    For Sor = S1 To Sx Step 1
        If Munka2.Range("a" & Sor).Value = Jel�ltT�rzs Then
            Tal�ltSor = Munka2.Range("a" & Sor).Row - 1
            End If
    Next
End If

Munka1.Select
Munka1.Range("w1").Value = Tal�ltSor
End Sub
