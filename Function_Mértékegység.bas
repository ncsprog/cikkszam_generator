Attribute VB_Name = "Function_Mértékegység"
Option Explicit

Sub Mértékegység()

Munka2.Select

Dim rngList As Range, JelöltSor As String
Set rngList = Munka2.Range("cq2:cq10")
AppCikkek.ComboBox15.List = rngList.Value
AppCikkek.ComboBox6.List = rngList.Value


End Sub
