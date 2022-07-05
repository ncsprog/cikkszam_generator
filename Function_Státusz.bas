Attribute VB_Name = "Function_Státusz"
Option Explicit

Sub Státusz()

Dim rngList As Range, JelöltSor As String
Set rngList = Munka2.Range("cw2:cw10")
AppCikkek.ComboBox9.List = rngList.Value

End Sub
