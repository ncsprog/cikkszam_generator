Attribute VB_Name = "Function_St�tusz"
Option Explicit

Sub St�tusz()

Dim rngList As Range, Jel�ltSor As String
Set rngList = Munka2.Range("cw2:cw10")
AppCikkek.ComboBox9.List = rngList.Value

End Sub
