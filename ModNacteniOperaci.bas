Attribute VB_Name = "ModNacteniOperaci"
Option Explicit

Public PocetZeta As Integer
Public PocetDeleni As Integer
Public PocetAlphaIDC As Integer
Public PocetDeleniTwist As Integer
Public PocetTwist As Integer
Public PocetITR As Integer
Public PocetTwisttube As Integer
Public PocetLisovani As Integer
Public PocetStul As Integer
Public PocetRBK As Integer
Public PocetBandaz As Integer
Public PocetZastrik As Integer
Public PocetSvar As Integer
Public PocetLumberg As Integer
Public PocetStocko As Integer
Public PocetRAMPF As Integer
Public PocetLaser As Integer
Public PocetMontaz As Integer
Public PocetOperaci As Integer
Public PocetListu As Integer
Public PocetOpDruhyList As Integer
Public PocetOpTretiList As Integer
Public PocetPracovniPostupy As Integer
Public PocetKaPr

Public ZetaProVyvojak() As Variant
Public Deleni() As Variant
Public AlphaIDC() As Variant
Public DeleniTwist() As Variant
Public Twist() As Variant
Public ITR() As Variant
Public Twisttube() As Variant
Public Lisovani() As Variant
Public Stul() As Variant
Public RBK() As Variant
Public Bandaz() As Variant
Public OpBandaz() As Variant
Public Zastrik() As Variant
Public OpZastrik() As Variant
Public Svar() As Variant
Public Lumberg() As Variant
Public Stocko() As Variant
Public RAMPF() As Variant
Public Laser() As Variant
Public Montaz() As Variant
Public Operace() As Variant
Public Ruzne() As Variant
Public OpDruhyList() As Variant
Public OpTretiList() As Variant
Public PracovniPostupy() As Variant
Public KartyProcesu() As Variant
Public CislaOpDeleni As String
Public CislaOpZeta As String
Public CislaOpAlphaIDC As String
Public CislaOpDeleniTwist As String
Public CislaOpTwisttube As String
Public CislaOpITRs As String
Public CislaOpTwist As String
Public CislaOpLisovani As String
Public CislaOpRBK As String
Public CislaOpStul As String
Public CislaOpBandaz As String
Public CislaOpZastrik As String
Public CislaOpSvar As String
Public CislaOpLumberg As String
Public CislaOpStocko As String
Public CislaOpRAMPF As String
Public CislaOpLaser As String
Public Posl As String

Sub ZobrazeniFormulare()
Attribute ZobrazeniFormulare.VB_Description = "Dotaz.Show"
Attribute ZobrazeniFormulare.VB_ProcData.VB_Invoke_Func = "V\n14"
  Dotaz.Show
End Sub

Sub NacteniDat()

Dim i As Integer
Dim J As Integer
Dim K As Boolean
Dim c_op As Integer
Dim PoradiOp As Integer
Dim index As Integer
Dim index2 As Integer
Dim index3 As Integer
Dim index4 As Integer
Dim index5 As Integer
Dim FindStroj As Range
Dim FindOperace As Range
Dim FindOznaceni As Range

PocetOperaci = ActiveWorkbook.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row - 1

Dim FindStrojNumber As Long
With ActiveWorkbook.Sheets(1)
    Set FindStroj = .Range("A1:O1").Find(What:="Stroj", LookIn:=xlValues)
End With
FindStrojNumber = FindStroj.Column

Dim FindOperaceNumber As Long
With ActiveWorkbook.Sheets(1)
    Set FindOperace = .Range("A1:O1").Find(What:="Prac. operace", LookIn:=xlValues)
End With
FindOperaceNumber = FindOperace.Column

Dim FindOznaceniNumber As Long
With ActiveWorkbook.Sheets(1)
    Set FindOznaceni = .Range("A1:O1").Find(What:="OznaËenÌ", LookIn:=xlValues)
End With
FindOznaceniNumber = FindOznaceni.Column

'Po zmÏnÏ nebo p¯id·nÌ pracoviöù je nutnÈ upravit Sub NacteniDat!
Dim PracovisteExcel As Application
Dim Pracoviste As Workbook
Set PracovisteExcel = New Application
PracovisteExcel.Visible = False
PracovisteExcel.Workbooks.Open ("P:\TPV\NOV¡ SLOéKA TPV\PROJEKTY\PPAP\V˝vojovÈ diagramy\KES\PracoviötÏ.xlsx")


For c_op = 2 To PocetOperaci + 1
  i = 2
  K = False
  Do While K = False And i < 94
    If i = 93 Then PocetMontaz = PocetMontaz + 1
    If ActiveWorkbook.Sheets(1).Cells(c_op, FindStrojNumber) = PracovisteExcel.Sheets("PracoviötÏ").Cells(i, 1) Then
      K = True
      If i > 1 And i < 21 Then PocetZeta = PocetZeta + 1
      If i > 20 And i < 46 Then PocetDeleni = PocetDeleni + 1
      If i > 45 And i < 51 Then PocetAlphaIDC = PocetAlphaIDC + 1
      If i > 50 And i < 53 Then PocetDeleniTwist = PocetDeleniTwist + 1
      If i > 52 And i < 55 Then PocetTwist = PocetTwist + 1
      If i > 54 And i < 57 Then PocetITR = PocetITR + 1
      If i = 57 Then PocetTwisttube = PocetTwisttube + 1
      If i > 57 And i < 71 Then PocetLisovani = PocetLisovani + 1
      If i = 71 Then PocetStul = PocetStul + 1
      If i = 72 Then PocetRBK = PocetRBK + 1
      If i > 72 And i < 82 Then PocetBandaz = PocetBandaz + 1
      If i > 81 And i < 87 Then PocetZastrik = PocetZastrik + 1
      If i = 87 Then PocetSvar = PocetSvar + 1
      If i > 87 And i < 90 Then PocetLumberg = PocetLumberg + 1
      If i = 90 Then PocetStocko = PocetStocko + 1
      If i = 91 Then PocetRAMPF = PocetRAMPF + 1
      If i = 92 Then PocetLaser = PocetLaser + 1
    End If
    i = i + 1
  Loop
Next c_op

PocetPracovniPostupy = PocetTwist + PocetStul + PocetRBK + PocetBandaz + PocetZastrik + PocetSvar + PocetLumberg + PocetStocko + PocetRAMPF + PocetLaser + PocetMontaz
If PocetPracovniPostupy > 0 Then ReDim PracovniPostupy(1 To PocetPracovniPostupy, 1 To 4)

PocetKaPr = PocetBandaz + PocetMontaz + PocetLaser
If PocetKaPr > 0 Then ReDim KartyProcesu(1 To PocetKaPr)

ReDim Operace(1 To PocetOperaci, 1 To 2)
If PocetMontaz > 0 Then ReDim Montaz(1 To PocetMontaz, 1 To 3)
If PocetBandaz > 0 Then ReDim OpBandaz(1 To PocetBandaz)
If PocetZastrik > 0 Then ReDim OpZastrik(1 To PocetZastrik)
If PocetRAMPF > 0 Then ReDim OpRAMPF(1 To PocetRAMPF)

PoradiOp = 1
index2 = 1
index3 = 1
index4 = 1
index5 = 1


For c_op = 2 To PocetOperaci + 1
  i = 2
  K = False
  Do While K = False And i < 94
    If i = 93 Then
      PoradiOp = PoradiOp + 1
      If c_op = PocetOperaci + 1 Then Posl = ActiveWorkbook.Sheets(1).Cells(c_op, FindOznaceniNumber)
      PracovniPostupy(index5, 2) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
      PracovniPostupy(index5, 3) = ActiveWorkbook.Sheets(1).Cells(c_op, FindStrojNumber)
      PracovniPostupy(index5, 4) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOznaceniNumber)
      PracovniPostupy(index5, 1) = "Mont·û"
      index5 = index5 + 1
      KartyProcesu(index4) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOznaceniNumber)
      index4 = index4 + 1
    End If
    If ActiveWorkbook.Sheets(1).Cells(c_op, FindStrojNumber) = PracovisteExcel.Sheets("PracoviötÏ").Cells(i, 1) Then
      K = True
      If i > 1 And i < 21 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "OsazenÌ konektoru"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "OsazenÌ konektoru"
      End If
      If i > 20 And i < 46 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "DÏlenÌ vodiËe + kontakt(y)"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "DÏlenÌ vodiËe + kontakt(y)"
      End If
      If i > 45 And i < 51 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "OsazenÌ konektoru STOCKO"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "OsazenÌ konektoru STOCKO"
      End If
      If i > 50 And i < 53 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "DÏlenÌ vodiËe + twist"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "DÏlenÌ vodiËe + twist"
      End If
      If i > 52 And i < 55 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "Twistov·nÌ vodiË˘"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "Twistov·nÌ vodiË˘"
        PracovniPostupy(index5, 1) = "Twistov·nÌ vodiË˘"
        index5 = index5 + 1
      End If
      If i > 54 And i < 57 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "DÏlenÌ ITRs"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "DÏlenÌ ITRs"
      End If
      If i = 57 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "DÏlenÌ twisttube"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "DÏlenÌ twisttube"
      End If
      If i > 57 And i < 71 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "Lisov·nÌ kontaktu"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "Lisov·nÌ kontaktu"
      End If
      If i = 71 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "Mont·û podskupiny"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "Mont·û podskupiny"
        PracovniPostupy(index5, 1) = "Mont·û podskupiny"
        index5 = index5 + 1
      End If
      If i = 72 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "Zafouk·nÌ ITRs"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "Zafouk·nÌ ITRs"
        PracovniPostupy(index5, 1) = "Zafouk·nÌ ITRs"
        index5 = index5 + 1
      End If
      If i > 72 And i < 82 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "StrojnÌ band·û"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "StrojnÌ band·û"
        OpBandaz(index2) = ActiveWorkbook.Sheets(1).Cells(c_op, FindStrojNumber)
        index2 = index2 + 1
        PracovniPostupy(index5, 1) = "StrojnÌ band·û"
        index5 = index5 + 1
        KartyProcesu(index4) = "StrojnÌ band·û"
        index4 = index4 + 1
      End If
      If i > 81 And i < 87 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "NÌzkotlak˝ z·st¯ik"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "NÌzkotlak˝ z·st¯ik"
        OpZastrik(index3) = ActiveWorkbook.Sheets(1).Cells(c_op, FindStrojNumber)
        index3 = index3 + 1
        PracovniPostupy(index5, 1) = "NÌzkotlak˝ z·st¯ik"
        index5 = index5 + 1
      End If
      If i = 87 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "Svar vodiË˘"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "Svar vodiË˘"
        PracovniPostupy(index5, 1) = "Svar vodiË˘"
        index5 = index5 + 1
      End If
      If i > 87 And i < 90 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "Osazov·nÌ konektoru LUMBERG"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "Osazov·nÌ konektoru LUMBERG"
        PracovniPostupy(index5, 1) = "Osazov·nÌ konektoru LUMBERG"
        index5 = index5 + 1
      End If
      If i = 90 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "Osazov·nÌ konektoru STOCKO"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "Osazov·nÌ konektoru STOCKO"
        PracovniPostupy(index5, 1) = "Osazov·nÌ konektoru STOCKO"
        index5 = index5 + 1
      End If
      If i = 91 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "RAMPF"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "RAMPF"
        PracovniPostupy(index5, 1) = "RAMPF"
        index5 = index5 + 1
      End If
      If i = 92 Then
        Operace(PoradiOp, 1) = ActiveWorkbook.Sheets(1).Cells(c_op, FindOperaceNumber)
        Operace(PoradiOp, 2) = "LaserovÈ znaËenÌ"
        PoradiOp = PoradiOp + 1
        If c_op = PocetOperaci + 1 Then Posl = "LaserovÈ znaËenÌ"
        PracovniPostupy(index5, 1) = "LaserovÈ znaËenÌ"
        index5 = index5 + 1
        KartyProcesu(index4) = "LaserovÈ znaËenÌ"
        index4 = index4 + 1
      End If
    End If
    i = i + 1
  Loop
Next c_op

Application.DisplayAlerts = False
PracovisteExcel.Workbooks.Close
Application.DisplayAlerts = True

End Sub

Sub ZapisOperaci(Pocet As Integer, Nazev As String, CislaOp As String, Poradi As Variant)

Dim c As Integer
Dim i As Integer
Dim d As Integer

If Pocet = 0 Then Exit Sub

ReDim Poradi(1 To Pocet)

Dim PoradiOp As Integer

c = 1
For PoradiOp = 1 To PocetOperaci
  If Operace(PoradiOp, 2) = Nazev Then
    Poradi(c) = PoradiOp
    c = c + 1
  End If
Next PoradiOp

d = 1
If Pocet = 1 Then
  CislaOp = "OP. " & Operace(Poradi(1), 1)
End If
For i = 1 To Pocet - 1
  If i = d Then
    If Poradi(1) + 1 <> Poradi(2) Then
      CislaOp = "OP. " & Operace(Poradi(1), 1) & ", "
    ElseIf Poradi(i) + 1 <> Poradi(i + 1) Then
      CislaOp = CislaOp + "OP. " & Operace(Poradi(i + 1 - d), 1) & " - " & Operace(Poradi(i), 1) & ", "
      d = 1
    Else
      d = d + 1
    End If
  ElseIf i > 1 And i <> d Then
    If Poradi(i) + 1 <> Poradi(i + 1) And Poradi(i) - 1 <> Poradi(i - 1) Then
      CislaOp = CislaOp + "OP. " & Operace(Poradi(i), 1) & ", "
      d = 1
    ElseIf Poradi(i) + 1 <> Poradi(i + 1) Then
      CislaOp = CislaOp + "OP. " & Operace(Poradi(i + 1 - d), 1) & " - " & Operace(Poradi(i), 1) & ", "
      d = 1
    Else
      d = d + 1
    End If
  End If
  If i = Pocet - 1 Then
    If d > 1 Then
      CislaOp = CislaOp + "OP. " & Operace(Poradi(i + 2 - d), 1) & " - " & Operace(Poradi(i + 1), 1)
    Else
      CislaOp = CislaOp + "OP. " & Operace(Poradi(i + 1), 1)
    End If
  End If
Next i

End Sub

Sub VypocetPoctuListu()

PocetListu = 2

If PocetDeleni > 0 Or PocetZeta > 0 Or PocetDeleniTwist > 0 Or PocetAlphaIDC > 0 Or PocetTwisttube > 0 Or PocetITR > 0 Then PocetListu = PocetListu + 1
If PocetTwist > 0 Or PocetLisovani > 0 Or PocetSvar > 0 Or PocetStocko > 0 Or PocetLumberg > 0 Then
  PocetListu = PocetListu + 1
End If
If PocetRBK > 0 Then PocetListu = PocetListu + 1
If PocetStul > 0 Then PocetListu = PocetListu + 1
If PocetBandaz > 0 Then PocetListu = PocetListu + 1
If PocetZastrik > 0 Then PocetListu = PocetListu + 1
If PocetRAMPF > 0 Then PocetListu = PocetListu + 1
If PocetLaser > 0 Then PocetListu = PocetListu + 1
If PocetMontaz > 0 Then PocetListu = PocetListu + PocetMontaz

End Sub

Sub VypocetPoctuOperaci()

Dim i As Integer
Dim a As Boolean
Dim b As Boolean
Dim c As Boolean
Dim d As Boolean
Dim e As Boolean
Dim f As Boolean
Dim K As Integer

If PocetListu > 2 Then
  If PocetDeleni > 0 Then PocetOpDruhyList = PocetOpDruhyList + 1
  If PocetZeta > 0 Then PocetOpDruhyList = PocetOpDruhyList + 1
  If PocetAlphaIDC > 0 Then PocetOpDruhyList = PocetOpDruhyList + 1
  If PocetTwisttube > 0 Then PocetOpDruhyList = PocetOpDruhyList + 1
  If PocetDeleniTwist > 0 Then PocetOpDruhyList = PocetOpDruhyList + 1
  If PocetITR > 0 Then PocetOpDruhyList = PocetOpDruhyList + 1
End If

If PocetOpDruhyList > 0 Then
  K = 1
  ReDim OpDruhyList(1 To PocetOpDruhyList)
  
  For i = 1 To PocetOperaci
    If Operace(i, 2) = "OsazenÌ konektoru" And a = False Then
      OpDruhyList(K) = "OsazenÌ konektoru"
      a = True
      K = K + 1
    ElseIf Operace(i, 2) = "DÏlenÌ vodiËe + kontakt(y)" And b = False Then
      OpDruhyList(K) = "DÏlenÌ vodiËe + kontakt(y)"
      b = True
      K = K + 1
    ElseIf Operace(i, 2) = "OsazenÌ konektoru STOCKO" And c = False Then
      OpDruhyList(K) = "OsazenÌ konektoru STOCKO"
      c = True
      K = K + 1
    ElseIf Operace(i, 2) = "DÏlenÌ vodiËe + twist" And d = False Then
      OpDruhyList(K) = "DÏlenÌ vodiËe + twist"
      d = True
      K = K + 1
    ElseIf Operace(i, 2) = "DÏlenÌ ITRs" And e = False Then
      OpDruhyList(K) = "DÏlenÌ ITRs"
      e = True
      K = K + 1
    ElseIf Operace(i, 2) = "DÏlenÌ twisttube" And f = False Then
      OpDruhyList(K) = "DÏlenÌ twisttube"
      f = True
      K = K + 1
    End If
  Next i
End If
  
  
If PocetListu > 3 Then
  If PocetTwist > 0 Then PocetOpTretiList = PocetOpTretiList + 1
  If PocetLisovani > 0 Then PocetOpTretiList = PocetOpTretiList + 1
  If PocetSvar > 0 Then PocetOpTretiList = PocetOpTretiList + 1
  If PocetStocko > 0 Then PocetOpTretiList = PocetOpTretiList + 1
  If PocetLumberg > 0 Then PocetOpTretiList = PocetOpTretiList + 1
End If

a = False
b = False
c = False
d = False
e = False
K = 1
If PocetOpTretiList > 0 Then
  
  ReDim OpTretiList(1 To PocetOpTretiList)
  
  For i = 1 To PocetOperaci
    If Operace(i, 2) = "Twistov·nÌ vodiË˘" And a = False Then
      OpTretiList(K) = "Twistov·nÌ vodiË˘"
      a = True
      K = K + 1
    ElseIf Operace(i, 2) = "Lisov·nÌ kontaktu" And b = False Then
      OpTretiList(K) = "Lisov·nÌ kontaktu"
      b = True
      K = K + 1
    ElseIf Operace(i, 2) = "Svar vodiË˘" And c = False Then
      OpTretiList(K) = "Svar vodiË˘"
      c = True
      K = K + 1
    ElseIf Operace(i, 2) = "Osazov·nÌ konektoru STOCKO" And d = False Then
      OpTretiList(K) = "Osazov·nÌ konektoru STOCKO"
      d = True
      K = K + 1
    ElseIf Operace(i, 2) = "Osazov·nÌ konektoru LUMBERG" And e = False Then
      OpTretiList(K) = "Osazov·nÌ konektoru LUMBERG"
      e = True
      K = K + 1
    End If
  Next i
End If

End Sub
  