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
Option Explicit

Private Sub CommandButton1_Click()
'Ellenõrzés
Mentés
Cikkszám_1
Cikkszám_2
Cikkszám_3
Cikkszám_4
IDgenerálás
ID_generálás
End Sub

Private Sub CommandButton2_Click()
Ûrlap_Törlés
End Sub

Private Sub UserForm_Initialize()
AppCikkek.TextBox9.Value = Date
'Utolsó_Cikkszám
Relevancia
Cikk_Kategória
End Sub
Sub ComboBox2_Click()
Cikktörzs
End Sub
Sub ComboBox3_Click()
Cikkosztály
End Sub
