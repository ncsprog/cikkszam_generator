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
'Ellen�rz�s
Ment�s
Cikksz�m_1
Cikksz�m_2
Cikksz�m_3
Cikksz�m_4
IDgener�l�s
ID_gener�l�s
End Sub

Private Sub CommandButton2_Click()
�rlap_T�rl�s
End Sub

Private Sub UserForm_Initialize()
AppCikkek.TextBox9.Value = Date
'Utols�_Cikksz�m
Relevancia
Cikk_Kateg�ria
End Sub
Sub ComboBox2_Click()
Cikkt�rzs
End Sub
Sub ComboBox3_Click()
Cikkoszt�ly
End Sub
