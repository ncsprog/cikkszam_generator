Attribute VB_Name = "Function_Cikktörzs"
Option Explicit

Sub Cikktörzs()
' - kész, lefut jó
Dim rngList As Range, JelöltSor As String
Set rngList = Munka2.Range("a2:a10")
AppCikkek.ComboBox12.List = rngList.Value

End Sub
