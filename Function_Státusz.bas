Attribute VB_Name = "Function_Státusz"
Option Explicit

Sub Státusz()

Munka2.Select

Dim rngList As Range, JelöltSor As String
Set rngList = Munka2.Range("cw2:cw10")
AppCikkek.ComboBox9.List = rngList.Value
AppCikkek.ComboBox5.List = rngList.Value

End Sub
