Attribute VB_Name = "Function_El�tag"
Option Explicit

Sub El�tag()

Munka2.Select

Dim rngList As Range, Jel�ltSor As String
Set rngList = Munka2.Range("co2:co10")
AppCikkek.ComboBox14.List = rngList.Value
AppCikkek.ComboBox7.List = rngList.Value

End Sub
