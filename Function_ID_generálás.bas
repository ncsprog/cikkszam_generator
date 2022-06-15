Attribute VB_Name = "Function_ID_generálás"
Option Explicit

Sub ID_generálás()

Dim most As Date
most = Now()

Sheets("adatok").Select
Columns("o:o").Select
Selection.End(xlDown).Select
Dim ID_nr As Long
ID_nr = ActiveCell + 1
Dim ID_rw As Long
ID_rw = ActiveCell.Row + 1
Dim ID_oszlop As String
ID_oszlop = "o"
Dim ID_koord As String
ID_koord = ID_oszlop & ID_rw
Range(ID_koord) = most


End Sub
