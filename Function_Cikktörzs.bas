Attribute VB_Name = "Function_Cikkt�rzs"
Option Explicit

Sub Cikkt�rzs()
' - k�sz, lefut j�

Munka2.Select

Dim rngList As Range, Jel�ltSor As String
Set rngList = Munka2.Range("a2:a10")
AppCikkek.ComboBox12.List = rngList.Value
AppCikkek.ComboBox2.List = rngList.Value

End Sub
