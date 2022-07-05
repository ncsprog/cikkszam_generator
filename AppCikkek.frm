VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AppCikkek 
   Caption         =   "Cikkek"
   ClientHeight    =   12420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   OleObjectBlob   =   "AppCikkek.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "AppCikkek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox11_Change()
AppCikkek.ComboBox10.Clear
Cikkfaj
End Sub

Private Sub ComboBox12_Change()
AppCikkek.ComboBox10.Clear
AppCikkek.ComboBox11.Clear
Cikkosztály
End Sub

Private Sub ComboBox2_Change()
Cikkosztály_lista

End Sub

Private Sub ComboBox3_Change()

Cikkfaj_lista

End Sub

Private Sub CommandButton1_Click()
ID_generálás
IDgenerálás

Cikkszám_1
Cikkszám_2
Cikkszám_3
Cikkszám_4
Mentés
Rekord_Visszaadás

End Sub
Private Sub CommandButton3_Click()  ' - Beállítások > Rögzítés
'
Elõtag_Csere    ' - ok
'
Relevancia_Csere    ' - ok
'
Mértékegység_Csere    ' - ok
'
Státusz_Csere    ' - ok
'
Cikktörzs_Csere    ' - ok
'
Cikkosztály_Csere    ' - ok
'
Cikkfaj_Csere   ' - ok

AppCikkek.ComboBox11.Clear
AppCikkek.ComboBox14.Clear
AppCikkek.ComboBox12.Clear
AppCikkek.ComboBox9.Clear
AppCikkek.ComboBox15.Clear
AppCikkek.ComboBox13.Clear
AppCikkek.ComboBox10.Clear
AppCikkek.TextBox20 = ""

UserForm_Initialize

End Sub


Private Sub CommandButton4_Click()

AppCikkek.ComboBox11.Clear
AppCikkek.ComboBox14.Clear
AppCikkek.ComboBox12.Clear
AppCikkek.ComboBox9.Clear
AppCikkek.ComboBox15.Clear
AppCikkek.ComboBox13.Clear
AppCikkek.ComboBox10.Clear
UserForm_Initialize

End Sub

Private Sub UserForm_Initialize()
Fontos
Elõtag
Relevancia
Mértékegység
Státusz
Cikktörzs


End Sub

