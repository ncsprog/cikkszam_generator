Attribute VB_Name = "Function_Cikkfaj_lista"
Option Explicit

Sub Cikkfaj_lista()

Munka2.Select

Dim JelöltOsztály, C1 As Integer, Cx As Integer, Collumns As Integer, TaláltOsztály As String, rngList As Range

JelöltOsztály = AppCikkek.ComboBox3.Value
C1 = Munka2.Range("k1").Column
Cx = Munka2.Range("cm1").Column

For Collumns = C1 To Cx Step 1
TaláltOsztály = Munka2.Cells(1, Collumns).Value
    If TaláltOsztály = JelöltOsztály Then
        Set rngList = Munka2.Range(Cells(2, Collumns), Cells(10, Collumns))
        AppCikkek.ComboBox4.List = rngList.Value
    End If
Next

End Sub
