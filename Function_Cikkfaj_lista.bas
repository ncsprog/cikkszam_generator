Attribute VB_Name = "Function_Cikkfaj_lista"
Option Explicit

Sub Cikkfaj_lista()

Munka2.Select

Dim Jel�ltOszt�ly, C1 As Integer, Cx As Integer, Collumns As Integer, Tal�ltOszt�ly As String, rngList As Range

Jel�ltOszt�ly = AppCikkek.ComboBox3.Value
C1 = Munka2.Range("k1").Column
Cx = Munka2.Range("cm1").Column

For Collumns = C1 To Cx Step 1
Tal�ltOszt�ly = Munka2.Cells(1, Collumns).Value
    If Tal�ltOszt�ly = Jel�ltOszt�ly Then
        Set rngList = Munka2.Range(Cells(2, Collumns), Cells(10, Collumns))
        AppCikkek.ComboBox4.List = rngList.Value
    End If
Next

End Sub
