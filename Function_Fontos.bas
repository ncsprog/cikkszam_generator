Attribute VB_Name = "Function_Fontos"
Option Explicit

Sub Fontos()

Munka2.Select

' - kész, lefut jó

AppCikkek.TextBox21.Value = Munka2.Range("cs2").Value
AppCikkek.TextBox22.Value = Munka2.Range("cy2").Value

End Sub
