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
Cikkoszt�ly
End Sub

Private Sub CommandButton1_Click()
ID_gener�l�s
IDgener�l�s
Ment�s
End Sub
Private Sub CommandButton3_Click()  ' - Be�ll�t�sok > R�gz�t�s
'
El�tag_Csere    ' - ok
'
Relevancia_Csere    ' - ok
'
M�rt�kegys�g_Csere    ' - ok
'
St�tusz_Csere    ' - ok
'
Cikkt�rzs_Csere    ' - ok
'
Cikkoszt�ly_Csere    ' - ok
'
Cikkfaj_Csere

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
El�tag
Relevancia
M�rt�kegys�g
St�tusz
Cikkt�rzs

End Sub

