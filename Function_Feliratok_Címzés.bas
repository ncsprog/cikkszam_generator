Attribute VB_Name = "Function_Feliratok_Címzés"
Option Explicit

Sub Feliratok_Címzés()

Munka2.Select

AppCikkek.Label2.Caption = Munka2.Range("dc3").Value
AppCikkek.Label3.Caption = Munka2.Range("dc4").Value
AppCikkek.Label4.Caption = Munka2.Range("dc5").Value
AppCikkek.Label5.Caption = Munka2.Range("dc6").Value
AppCikkek.Label6.Caption = Munka2.Range("dc7").Value
AppCikkek.Label7.Caption = Munka2.Range("dc10").Value
AppCikkek.Label8.Caption = Munka2.Range("dc9").Value
AppCikkek.Label9.Caption = Munka2.Range("dc8").Value
AppCikkek.Label10.Caption = Munka2.Range("dc16").Value
AppCikkek.Label11.Caption = Munka2.Range("dc13").Value
AppCikkek.Label12.Caption = Munka2.Range("dc11").Value
AppCikkek.Label13.Caption = Munka2.Range("dc17").Value
'AppCikkek.Label14.Caption = Munka2.Range("").Value
'AppCikkek.Label15.Caption = Munka2.Range("").Value
AppCikkek.Label16.Caption = Munka2.Range("dc22").Value
AppCikkek.Label17.Caption = Munka2.Range("dc18").Value
AppCikkek.Label18.Caption = Munka2.Range("dc19").Value
AppCikkek.Label19.Caption = Munka2.Range("dc20").Value
AppCikkek.Label20.Caption = Munka2.Range("dc21").Value
AppCikkek.Label21.Caption = Munka2.Range("dc2").Value
AppCikkek.Label22.Caption = Munka2.Range("dc14").Value
AppCikkek.Label23.Caption = Munka2.Range("dc15").Value
AppCikkek.Label24.Caption = Munka2.Range("dc7").Value
AppCikkek.Label25.Caption = Munka2.Range("dc6").Value
AppCikkek.Label26.Caption = Munka2.Range("dc5").Value
AppCikkek.Label27.Caption = Munka2.Range("dc4").Value
AppCikkek.Label28.Caption = Munka2.Range("dc3").Value
AppCikkek.Label29.Caption = Munka2.Range("dc2").Value
AppCikkek.Label30.Caption = Munka2.Range("dc22").Value
AppCikkek.Label31.Caption = Munka2.Range("dc30").Value
AppCikkek.Label32.Caption = Munka2.Range("dc7").Value
AppCikkek.Label33.Caption = Munka2.Range("dc8").Value
AppCikkek.Label34.Caption = Munka2.Range("dc9").Value
AppCikkek.Label35.Caption = Munka2.Range("dc10").Value
AppCikkek.Label36.Caption = Munka2.Range("dc11").Value
AppCikkek.Label37.Caption = Munka2.Range("dc17").Value
AppCikkek.Label38.Caption = Munka2.Range("dc29").Value
'AppCikkek.Label39.Caption = Munka2.Range("").Value
AppCikkek.Label40.Caption = Munka2.Range("dc8").Value

AppCikkek.Frame9.Caption = Munka2.Range("dc27").Value
AppCikkek.Frame10.Caption = Munka2.Range("dc12").Value
''AppCikkek.Frame11.Caption = Munka2.Range("").Value
''AppCikkek.Frame12.Caption = Munka2.Range("").Value
'
AppCikkek.CommandButton1.Caption = Munka2.Range("dc24").Value
AppCikkek.CommandButton2.Caption = Munka2.Range("dc23").Value
AppCikkek.CommandButton3.Caption = Munka2.Range("dc24").Value
AppCikkek.CommandButton4.Caption = Munka2.Range("dc26").Value
AppCikkek.CommandButton5.Caption = Munka2.Range("dc28").Value
AppCikkek.CommandButton6.Caption = Munka2.Range("dc28").Value

End Sub
