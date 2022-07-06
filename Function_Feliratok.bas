Attribute VB_Name = "Function_Feliratok"
Option Explicit

Sub Feliratok()

Munka2.Select

Dim rngList As Range, JelöltSor As String
Set rngList = Munka2.Range("dc2:dc30")
AppCikkek.ComboBox16.List = rngList.Value

End Sub
