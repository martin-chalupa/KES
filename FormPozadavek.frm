VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPozadavek 
   Caption         =   "Vytvoøit požadavek na validaci"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3840
   OleObjectBlob   =   "FormPozadavek.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPozadavek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ButPozadavek_Click()
  If Len(TxtCisloSvazku.Text) <> 13 Then
    MsgBox "Zadej èíslo svazku ve formátu KES XXX.XX.XXX.XX!"
    Exit Sub
  End If
  Call ModPozadavekNaValidaci.PozadavekNaValidaci
  If Bunka Is Nothing Then Exit Sub
  ThisWorkbook.Sheets("DATA1").Activate
  If ValidaceX = False And ValidaceY = False Then
    MsgBox "Pro svazek " & CisloSvazku & " jsou všechny kombinace zvalidovány."
    Exit Sub
  Else
    MsgBox "Doplò výrobce materiálu a konektor!" & vbNewLine & "Zkontroluj normu a zda namísto èísla tìsnìní není uvedeno èíslo konektoru!"
  End If
  Unload Me
  End
End Sub

Private Sub ButStorno_Click()
  Unload Me
  End
End Sub
