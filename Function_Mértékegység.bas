Attribute VB_Name = "Function_M�rt�kegys�g"
Option Explicit

Sub M�rt�kegys�g()

Munka2.Select

Dim rngList As Range, Jel�ltSor As String
Set rngList = Munka2.Range("cq2:cq10")
AppCikkek.ComboBox15.List = rngList.Value
AppCikkek.ComboBox6.List = rngList.Value


End Sub
