Attribute VB_Name = "Function_Elõtag"
Option Explicit

Sub Elõtag()

Dim rngList As Range, JelöltSor As String
Set rngList = Munka2.Range("co2:co10")
AppCikkek.ComboBox14.List = rngList.Value
End Sub
