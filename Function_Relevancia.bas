Attribute VB_Name = "Function_Relevancia"
Option Explicit

Sub Relevancia()

Dim rngList As Range, Jel�ltSor As String
Set rngList = Munka2.Range("cu2:cu10")
AppCikkek.ComboBox13.List = rngList.Value

End Sub
