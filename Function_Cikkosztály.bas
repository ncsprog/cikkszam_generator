Attribute VB_Name = "Function_Cikkoszt�ly"
Option Explicit

Sub Cikkoszt�ly()

Dim Jel�ltT�rzs, C1 As Integer, Cx As Integer, Collumns As Integer, Tal�ltT�rzs As String, rngList As Range

Jel�ltT�rzs = AppCikkek.ComboBox12.Value
C1 = Munka2.Range("b1").Column
Cx = Munka2.Range("j1").Column

For Collumns = C1 To Cx Step 1
Tal�ltT�rzs = Munka2.Cells(1, Collumns).Value
    If Tal�ltT�rzs = Jel�ltT�rzs Then
        Set rngList = Munka2.Range(Cells(2, Collumns), Cells(10, Collumns))
        AppCikkek.ComboBox11.List = rngList.Value
    End If
Next

End Sub
