Attribute VB_Name = "ModVyvojak"
Option Explicit
Public e As Integer

Sub PrvniList()

ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = "VD1"
ActiveWorkbook.Sheets("VD1").Select

ActiveSheet.Range("C2:G2").Merge
ActiveSheet.Range("C3:G3").Merge
ActiveSheet.Range("C4:E4").Merge
ActiveSheet.Range("B8:G8").Merge
ActiveSheet.Range("B9:G9").Merge
ActiveSheet.Range("A1").ColumnWidth = 30
ActiveSheet.Range("G1").ColumnWidth = 15
ActiveSheet.Range("C2:G2").Value2 = ("VÝVOJOVÝ DIAGRAM" & vbNewLine & "FLOW CHART")
ActiveSheet.Range("C2:G2").Characters(Start:=0, Length:=28).Font.Bold = True
ActiveSheet.Range("C2:G2").Characters(Start:=0, Length:=16).Font.Size = 18
ActiveSheet.Range("C2:G2").Characters(Start:=17, Length:=26).Font.Italic = True
ActiveSheet.Range("C2:G2").HorizontalAlignment = xlCenter
ActiveSheet.Range("C2:G2").VerticalAlignment = xlCenter
ActiveSheet.Range("B2").RowHeight = 55
ActiveSheet.Range("B1").ColumnWidth = 30
ActiveSheet.Range("C1").ColumnWidth = 19
ActiveSheet.Range("B3:B5").RowHeight = 16.1
ActiveSheet.Range("B3").Value2 = ("Èíslo výrobku / Product No.:")
ActiveSheet.Range("B3").Characters(Start:=16, Length:=26).Font.Italic = True
ActiveSheet.Range("B3:H6").HorizontalAlignment = xlRight
ActiveSheet.Range("B4").Value2 = ("Název výrobku / Description:")
ActiveSheet.Range("B4").Characters(Start:=17, Length:=28).Font.Italic = True
ActiveSheet.Range("B5").Value2 = ("Èíslo výkresu / Drawing No.:")
ActiveSheet.Range("B5").Characters(Start:=17, Length:=27).Font.Italic = True
ActiveSheet.Range("D5").Value2 = ("Index:")
ActiveSheet.Range("D1").ColumnWidth = 6.5
ActiveSheet.Range("F4").Value2 = ("Vypracoval / Processed by:")
ActiveSheet.Range("F4").Characters(Start:=14, Length:=26).Font.Italic = True
ActiveSheet.Range("F1").ColumnWidth = 24
ActiveSheet.Range("F5").Value2 = ("Datum:")
ActiveSheet.Range("G5").NumberFormat = "d.m.yyyy"
ActiveSheet.Range("G5").Value2 = Date
ActiveSheet.Range("G5").HorizontalAlignment = xlHAlignLeft
ActiveSheet.Range("B8:G8").Value2 = "               Výrobek ve všech fázích pøepravovat manipulaèním vozíkem KES (pokud není uvedeno jinak)."
ActiveSheet.Range("B9:G9").Value2 = "               In all product phases the product to be transported by KES material handling cart (if not specified)."
ActiveSheet.Range("B9:G9").Characters.Font.Italic = True

ActiveSheet.Range("C3").Value = Dotaz.TxtCisloSvazku
ActiveSheet.Range("C3").HorizontalAlignment = xlHAlignLeft
ActiveSheet.Range("C4").Value = Dotaz.TxtVyrobek
ActiveSheet.Range("C4").HorizontalAlignment = xlHAlignLeft
ActiveSheet.Range("G4").Value = Dotaz.TxtJmeno
ActiveSheet.Range("G4").HorizontalAlignment = xlHAlignLeft
ActiveSheet.Range("G4").EntireColumn.AutoFit
ActiveSheet.Range("C5").Value = Dotaz.TxtCisloVykresu
ActiveSheet.Range("C5").HorizontalAlignment = xlHAlignLeft
ActiveSheet.Range("E5").Value = Dotaz.TxtIndex
ActiveSheet.Range("E5").HorizontalAlignment = xlHAlignLeft
ActiveSheet.Range("E5").EntireColumn.AutoFit

With ActiveSheet.Pictures.Insert("P:\TPV\NOVÁ SLOŽKA TPV\PROJEKTY\PPAP\Vývojové diagramy\KES\KES logo.jpg")
    .Left = 180
    .Top = 19
    .Width = 80
    .Height = 50
End With

ActiveSheet.Range("B2:F5").Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
ActiveSheet.Range("B2:F5").Select
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
ActiveSheet.Range("B2:F5").Select
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
ActiveSheet.Range("B2:F5").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = xlAutomatic
    End With
ActiveSheet.Range("B2:B5").Select
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
ActiveSheet.Range("B3:G3").Select
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
ActiveSheet.Range("B3:G3").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
ActiveSheet.Range("B4:G4").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
ActiveSheet.Range("D5").Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
ActiveSheet.Range("D5").Select
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
ActiveSheet.Range("F4:F5").Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
ActiveSheet.Range("F4:F5").Select
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

Call ModObjekty.Postup(250, 200, 120, 37)
Call Textbox(250, 200, 120, 37, "Pøíjem materiálu" & vbNewLine & "Receive of material", 17, 18, 26, True, True)
Call Rozhodnuti(215, 272, 190, 70)
Call Textbox(250, 290, 125, 37, "Je dodávka v poøádku?" & vbNewLine & "Is supplied material OK?", 22, 23, 28, True, True)
Call Postup(250, 377, 120, 37)
Call Textbox(250, 377, 120, 37, "Pøíjmový sklad XX" & vbNewLine & "Receiving stock XX", 17, 18, 26, True, True)
Call Postup(250, 449, 120, 37)
Call Textbox(250, 449, 120, 37, "Vstupní kontrola" & vbNewLine & "Incoming inspection", 17, 18, 26, True, True)
Call Rozhodnuti(215, 521, 190, 70)
Call Textbox(240, 542, 140, 37, "Vyhovuje materiál specifikacím?" & vbNewLine & "Does the material match the specification?", 32, 33, 42, True, True)
Call Postup(250, 626, 120, 37)
Call Textbox(250, 626, 120, 37, "Sklad MP" & vbNewLine & "Stock MP", 8, 9, 17, True, True)
Call Postup(250, 698, 120, 37)
Call Textbox(250, 698, 120, 37, "Sklad V1" & vbNewLine & "Stock V1", 8, 9, 17, True, True)
Call Dokument(530, 200, 100, 50)
Call Textbox(530, 200, 100, 30, "OS 7.5-06" & vbNewLine & "Dodací list" & vbNewLine & "Delivery note", 9, 22, 0, True, False)
Call Dokument(530, 280, 100, 50)
Call Textbox(530, 290, 100, 30, "OS 8.3-01", 9, 9, 0, False, False)
Call Dokument(530, 377, 100, 50)
Call Textbox(530, 377, 100, 30, "OS 7.5-06" & vbNewLine & "Doklad o pøíjmu zboží" & vbNewLine & "Stock receipt note", 35, 32, 0, True, False)
Call Dokument(500, 449, 155, 50)
Call Textbox(500, 449, 160, 30, "OS 8.1-01" & vbNewLine & "Kontrolní postupy pro vstupní pøejímku" & vbNewLine & "Controls plans for incoming inspection", 9, 22, 0, True, False)
Call Dokument(530, 532, 100, 50)
Call Textbox(530, 542, 100, 30, "OS 8.3-01", 9, 9, 0, False, False)
Call Dokument(530, 626, 100, 50)
Call Textbox(530, 626, 100, 30, "Ident. Lístek ´Pøíjem´" & vbNewLine & "Ident. ´Receipt´", 22, 22, 0, True, False)
Call Dokument(530, 698, 100, 50)
Call Textbox(530, 698, 100, 30, "Ident. Lístek ´Odbìr´" & vbNewLine & "Ident. ´Removal´", 22, 22, 0, True, False)
Call Kruh(297.5, 770, 25, 25)
Call Textbox(298, 773, 25, 25, "1", 1, 1, 1, False, False)
Call Sipka(310, 242, 310, 267)
Call Sipka(310, 347, 310, 372)
Call Sipka(310, 419, 310, 444)
Call Sipka(310, 491, 310, 516)
Call Sipka(310, 596, 310, 621)
Call Sipka(310, 668, 310, 693)
Call Sipka(410, 307, 525, 307)
Call Sipka(410, 556, 525, 556)
Call Sipka(310, 740, 310, 765)
Call Textbox(310, 340, 50, 30, "Ano" & vbNewLine & "Yes", 3, 4, 0, True, False)
Call Textbox(310, 589, 50, 30, "Ano" & vbNewLine & "Yes", 3, 4, 0, True, False)
Call Textbox(185, 590, 130, 30, "Paletový vozík / Pallet cart", 17, 17, 0, True, False)
Call Textbox(176, 601, 130, 30, "Vysokozdvižný vozík / Fork cart", 22, 22, 0, True, False)
Call Textbox(440, 275, 50, 30, "Ne" & vbNewLine & "No", 2, 3, 0, True, False)
Call Textbox(440, 525, 50, 30, "Ne" & vbNewLine & "No", 2, 3, 0, True, False)

Application.ScreenUpdating = True

ActiveSheet.Columns("A").EntireColumn.Delete

With ActiveSheet.PageSetup
 .Zoom = False
 .FitToPagesTall = 1
 .FitToPagesWide = 1
End With

Application.ScreenUpdating = False

End Sub

Sub DruhyList()
  Dim i As Integer
  If PocetDeleni < 1 And PocetZeta < 1 And PocetAlphaIDC < 1 And PocetTwisttube < 1 And PocetDeleniTwist < 1 And PocetITR < 1 Then Exit Sub
  
  ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = "VD2"
  ActiveWorkbook.Sheets("VD2").Select

  ActiveSheet.PageSetup.Orientation = xlLandscape
    
  Call ModObjekty.Kruh(401, 15, 25, 25)
  Call Textbox(401, 18, 25, 25, "1", 1, 1, 1, False, False)
  Call Cara(413, 45, 413, 65)
  
  Call ModNacteniOperaci.ZapisOperaci(PocetDeleni, "Dìlení vodièe + kontakt(y)", CislaOpDeleni, Deleni)
  Call ModNacteniOperaci.ZapisOperaci(PocetZeta, "Osazení konektoru", CislaOpZeta, ZetaProVyvojak)
  Call ModNacteniOperaci.ZapisOperaci(PocetAlphaIDC, "Osazení konektoru STOCKO", CislaOpAlphaIDC, AlphaIDC)
  Call ModNacteniOperaci.ZapisOperaci(PocetDeleniTwist, "Dìlení vodièe + twist", CislaOpDeleniTwist, DeleniTwist)
  Call ModNacteniOperaci.ZapisOperaci(PocetTwisttube, "Dìlení twisttube", CislaOpTwisttube, Twisttube)
  Call ModNacteniOperaci.ZapisOperaci(PocetITR, "Dìlení ITRs", CislaOpITRs, ITR)

  If PocetOpDruhyList = 1 Then
    Call ModObjekty.Sipka(413, 65, 413, 85)
  ElseIf PocetOpDruhyList = 2 Then
    Call Cara(345, 65, 480, 65)
    Call Sipka(345, 65, 345, 85)
    Call Sipka(480, 65, 480, 85)
    Call Cara(345, 285, 480, 285)
  ElseIf PocetOpDruhyList = 3 Then
    Call Cara(278, 65, 548, 65)
    Call Sipka(413, 65, 413, 85)
    Call Sipka(278, 65, 278, 85)
    Call Sipka(548, 65, 548, 85)
    Call Cara(278, 285, 548, 285)
  ElseIf PocetOpDruhyList = 4 Then
    Call Cara(210, 65, 615, 65)
    Call Sipka(210, 65, 210, 85)
    Call Sipka(345, 65, 345, 85)
    Call Sipka(480, 65, 480, 85)
    Call Sipka(615, 65, 615, 85)
    Call Cara(210, 285, 615, 285)
  ElseIf PocetOpDruhyList = 5 Then
    Call Cara(143, 65, 683, 65)
    Call Sipka(413, 65, 413, 85)
    Call Sipka(143, 65, 143, 85)
    Call Sipka(278, 65, 278, 85)
    Call Sipka(548, 65, 548, 85)
    Call Sipka(683, 65, 683, 85)
    Call Cara(143, 285, 683, 285)
  ElseIf PocetOpDruhyList = 6 Then
    Call Cara(75, 65, 750, 65)
    Call Sipka(75, 65, 75, 85)
    Call Sipka(210, 65, 210, 85)
    Call Sipka(345, 65, 345, 85)
    Call Sipka(480, 65, 480, 85)
    Call Sipka(615, 65, 615, 85)
    Call Sipka(750, 65, 750, 85)
    Call Cara(75, 285, 750, 285)
  End If

  For i = 1 To PocetOpDruhyList
    If PocetOpDruhyList = 1 Then
      Call DruhyListOp(0, 1)
    ElseIf PocetOpDruhyList = 2 Then
      If i = 1 Then Call DruhyListOp(-68, 1)
      If i = 2 Then Call DruhyListOp(68, 2)
    ElseIf PocetOpDruhyList = 3 Then
      If i = 1 Then Call DruhyListOp(-135, 1)
      If i = 2 Then Call DruhyListOp(0, 2)
      If i = 3 Then Call DruhyListOp(135, 3)
    ElseIf PocetOpDruhyList = 4 Then
      If i = 1 Then Call DruhyListOp(-203, 1)
      If i = 2 Then Call DruhyListOp(-68, 2)
      If i = 3 Then Call DruhyListOp(68, 3)
      If i = 4 Then Call DruhyListOp(203, 4)
    ElseIf PocetOpDruhyList = 5 Then
      If i = 1 Then Call DruhyListOp(-270, 1)
      If i = 2 Then Call DruhyListOp(-135, 2)
      If i = 3 Then Call DruhyListOp(0, 3)
      If i = 4 Then Call DruhyListOp(135, 4)
      If i = 5 Then Call DruhyListOp(270, 5)
    ElseIf PocetOpDruhyList = 6 Then
      If i = 1 Then Call DruhyListOp(-338, 1)
      If i = 2 Then Call DruhyListOp(-203, 2)
      If i = 3 Then Call DruhyListOp(-68, 3)
      If i = 4 Then Call DruhyListOp(68, 4)
      If i = 5 Then Call DruhyListOp(203, 5)
      If i = 6 Then Call DruhyListOp(338, 6)
    End If
  Next i

  Call Sipka(413, 285, 413, 305)
  Call Rozhodnuti(313, 310, 200, 80)
  Call Textbox(340, 335, 150, 70, "Vyhovuje materiál specifikacím?" & vbNewLine & "Does the material match the specification?", 32, 33, 42, True, True)
  Call Sipka(518, 350, 557, 350)
  Call Textbox(510, 320, 50, 50, "Ne" & vbNewLine & "No", 2, 3, 3, True, False)
  Call Sipka(413, 395, 413, 415)
  Call Textbox(410, 388, 50, 50, "Ano" & vbNewLine & "Yes", 2, 3, 3, True, False)
  Call Dokument(562, 325, 100, 50)
  Call Textbox(562, 340, 100, 30, "OS 8.3-01", 9, 9, 0, False, False)
  Call Postup(363, 420, 100, 30)
  Call Textbox(363, 420, 100, 30, "Sklad V1" & vbNewLine & "Stock V1", 8, 9, 17, True, True)
  Call Dokument(363, 450, 100, 35)
  Call Textbox(363, 450, 100, 35, "Prùvodní lístek" & vbNewLine & "Carriage note", 0, 16, 28, True, False)
  Call Sipka(413, 490, 413, 515)
  Call Kruh(401, 520, 25, 25)
  Call Textbox(401, 524, 25, 25, "2", 1, 1, 1, False, False)
  
  ActiveSheet.Columns("A").EntireColumn.Delete
    
  With ActiveSheet.PageSetup
   .Zoom = False
   .FitToPagesTall = 1
   .FitToPagesWide = 1
  End With
End Sub

Sub DruhyListOp(a As Double, i As Integer)
  If OpDruhyList(i) = "Osazení konektoru" Then
    Call Postup(348 + a, 90, 130, 65)
    If Posl = "Osazení konektoru" Then
      Call Dokument(348 + a, 155, 130, 125)
      Call Textbox(348 + a, 235, 130, 10, "Balicí list", 0, 0, 0, False, False)
      Call Textbox(348 + a, 245, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
      Call Cara(413 + a, 280, 413 + a, 285)
    Else
      Call Dokument(348 + a, 155, 130, 105)
      Call Cara(413 + a, 260, 413 + a, 285)
    End If
    Call TxtOperace(348 + a, 90, 130, 25, CislaOpZeta, Len(CislaOpZeta), 0, 0, False, True)
    Call Textbox(348 + a, 110, 130, 10, "Osazení konektoru", 17, 0, 0, False, True)
    Call Textbox(348 + a, 120, 130, 10, "Connector`s assembly", 20, 1, 19, True, True)
    Call Textbox(348 + a, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
    Call Textbox(348 + a, 140, 130, 10, "Komax Zeta", 0, 0, 0, False, False)
    Call Textbox(348 + a, 155, 130, 90, "Typový kontrolní postup" & vbNewLine & "000.05" & vbNewLine & "Type control plan 000.05", 0, 31, 27, True, False)
    Call Textbox(348 + a, 190, 130, 30, "Nástøihový plán" & vbNewLine & "Cutting plan", 0, 16, 12, True, False)
    Call Textbox(348 + a, 213, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
  ElseIf OpDruhyList(i) = "Dìlení vodièe + kontakt(y)" Then
    Call Postup(348 + a, 90, 130, 65)
    If Posl = "Dìlení vodièe + kontakt(y)" Then
      Call Dokument(348 + a, 155, 130, 125)
      Call Textbox(348 + a, 235, 130, 10, "Balicí list", 0, 0, 0, False, False)
      Call Textbox(348 + a, 245, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
      Call Cara(413 + a, 280, 413 + a, 285)
    Else
      Call Dokument(348 + a, 155, 130, 105)
      Call Cara(413 + a, 260, 413 + a, 285)
    End If
    Call TxtOperace(348 + a, 90, 130, 25, CislaOpDeleni, Len(CislaOpDeleni), 0, 0, False, True)
    Call Textbox(348 + a, 110, 130, 10, "Dìlení vodièe + kontakt(y)", 26, 0, 0, False, True)
    Call Textbox(348 + a, 120, 130, 10, "Cutting of wire + contact(s)", 28, 1, 28, True, True)
    Call Textbox(348 + a, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
    Call Textbox(348 + a, 140, 130, 10, "Komax Alpha, Gamma", 0, 0, 0, False, False)
    Call Textbox(348 + a, 155, 130, 90, "Typový kontrolní postup" & vbNewLine & "000.05" & vbNewLine & "Type control plan 000.05", 0, 31, 27, True, False)
    Call Textbox(348 + a, 190, 130, 30, "Nástøihový plán" & vbNewLine & "Cutting plan", 0, 16, 12, True, False)
    Call Textbox(348 + a, 213, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
  ElseIf OpDruhyList(i) = "Osazení konektoru STOCKO" Then
    Call Postup(348 + a, 90, 130, 65)
    If Posl = "Osazení konektoru STOCKO" Then
      Call Dokument(348 + a, 155, 130, 125)
      Call Textbox(348 + a, 235, 130, 10, "Balicí list", 0, 0, 0, False, False)
      Call Textbox(348 + a, 245, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
      Call Cara(413 + a, 280, 413 + a, 285)
    Else
      Call Dokument(348 + a, 155, 130, 105)
      Call Cara(413 + a, 260, 413 + a, 285)
    End If
    Call TxtOperace(348 + a, 90, 130, 25, CislaOpAlphaIDC, Len(CislaOpAlphaIDC), 0, 0, False, True)
    Call Textbox(348 + a, 110, 130, 10, "Osazení konektoru Stocko", 26, 0, 0, False, True)
    Call Textbox(348 + a, 120, 130, 10, "Stocko connector`s assembly", 27, 1, 27, True, True)
    Call Textbox(348 + a, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
    Call Textbox(348 + a, 140, 130, 10, "Komax Alpha IDC", 0, 0, 0, False, False)
    Call Textbox(348 + a, 155, 130, 90, "Typový kontrolní postup" & vbNewLine & "000.05" & vbNewLine & "Type control plan 000.05", 0, 31, 27, True, False)
    Call Textbox(348 + a, 190, 130, 30, "Nástøihový plán" & vbNewLine & "Cutting plan", 0, 16, 12, True, False)
    Call Textbox(348 + a, 213, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
  ElseIf OpDruhyList(i) = "Dìlení vodièe + twist" Then
    Call Postup(348 + a, 90, 130, 65)
    If Posl = "Dìlení vodièe + twist" Then
      Call Dokument(348 + a, 155, 130, 125)
      Call Textbox(348 + a, 235, 130, 10, "Balicí list", 0, 0, 0, False, False)
      Call Textbox(348 + a, 245, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
      Call Cara(413 + a, 280, 413 + a, 285)
    Else
      Call Dokument(348 + a, 155, 130, 105)
      Call Cara(413 + a, 260, 413 + a, 285)
    End If
    Call TxtOperace(348 + a, 90, 130, 25, CislaOpDeleniTwist, Len(CislaOpDeleniTwist), 0, 0, False, True)
    Call Textbox(348 + a, 110, 130, 10, "Dìlení vodièe + twist", 21, 0, 0, False, True)
    Call Textbox(348 + a, 120, 130, 10, "Cutting of wire + twist", 23, 1, 27, True, True)
    Call Textbox(348 + a, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
    Call Textbox(348 + a, 140, 130, 10, "Komax Alpha", 0, 0, 0, False, False)
    Call Textbox(348 + a, 155, 130, 90, "Typový kontrolní postup" & vbNewLine & "000.05" & vbNewLine & "Type control plan 000.05", 0, 31, 27, True, False)
    Call Textbox(348 + a, 190, 130, 30, "Nástøihový plán" & vbNewLine & "Cutting plan", 0, 16, 12, True, False)
    Call Textbox(348 + a, 213, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
  ElseIf OpDruhyList(i) = "Dìlení twisttube" Then
    Call Postup(348 + a, 90, 130, 65)
    Call Dokument(348 + a, 155, 130, 105)
    Call Cara(413 + a, 260, 413 + a, 285)
    Call TxtOperace(348 + a, 90, 130, 25, CislaOpTwisttube, Len(CislaOpTwisttube), 0, 0, False, True)
    Call Textbox(348 + a, 110, 130, 10, "Dìlení twisttube", 16, 0, 0, False, True)
    Call Textbox(348 + a, 120, 130, 10, "Cutting of twisttube", 20, 1, 20, True, True)
    Call Textbox(348 + a, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
    Call Textbox(348 + a, 140, 130, 10, "24-302", 0, 0, 0, False, False)
    Call Textbox(348 + a, 155, 130, 90, "Typový kontrolní postup" & vbNewLine & "000.15" & vbNewLine & "Type control plan 000.15", 0, 31, 27, True, False)
    Call Textbox(348 + a, 190, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
    If Posl = "Dìlení twisttube" Then
      Call Textbox(348 + a, 213, 130, 10, "Balicí list", 0, 0, 0, False, False)
      Call Textbox(348 + a, 223, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
    End If
  ElseIf OpDruhyList(i) = "Dìlení ITRs" Then
    Call Postup(348 + a, 90, 130, 65)
    Call Dokument(348 + a, 155, 130, 105)
    Call Cara(413 + a, 260, 413 + a, 285)
    Call TxtOperace(348 + a, 90, 130, 25, CislaOpITRs, Len(CislaOpITRs), 0, 0, False, True)
    Call Textbox(348 + a, 110, 130, 10, "Dìlení ITR(s)", 13, 0, 0, False, True)
    Call Textbox(348 + a, 120, 130, 10, "Cutting of insulation tube", 26, 1, 26, True, True)
    Call Textbox(348 + a, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
    Call Textbox(348 + a, 140, 130, 10, "24-301/24-303", 0, 0, 0, False, False)
    Call Textbox(348 + a, 155, 130, 90, "Typový kontrolní postup" & vbNewLine & "000.03" & vbNewLine & "Type control plan 000.03", 0, 31, 27, True, False)
    Call Textbox(348 + a, 190, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
    If Posl = "Dìlení ITRs" Then
      Call Textbox(348 + a, 213, 130, 10, "Balicí list", 0, 0, 0, False, False)
      Call Textbox(348 + a, 223, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
    End If
  End If
End Sub

Sub TretiList()
  Dim i As Integer
  If PocetTwist < 1 And PocetLisovani < 1 And PocetSvar < 1 And PocetStocko < 1 And PocetLumberg < 1 Then Exit Sub
  ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = "VD3"
  ActiveWorkbook.Sheets("VD3").Select

  ActiveSheet.PageSetup.Orientation = xlLandscape
  
  Call ModObjekty.Kruh(401, 15, 25, 25)
  Call Textbox(401, 18, 25, 25, "2", 1, 1, 1, False, False)
  Call Cara(413, 45, 413, 65)

  Call ModNacteniOperaci.ZapisOperaci(PocetTwist, "Twistování vodièù", CislaOpTwist, Twist)
  Call ModNacteniOperaci.ZapisOperaci(PocetLisovani, "Lisování kontaktu", CislaOpLisovani, Lisovani)
  Call ModNacteniOperaci.ZapisOperaci(PocetSvar, "Svar vodièù", CislaOpSvar, Svar)
  Call ModNacteniOperaci.ZapisOperaci(PocetStocko, "Osazování konektoru STOCKO", CislaOpStocko, Stocko)
  Call ModNacteniOperaci.ZapisOperaci(PocetLumberg, "Osazování konektoru LUMBERG", CislaOpLumberg, Lumberg)

  If PocetOpTretiList = 1 Then
    Call ModObjekty.Sipka(413, 65, 413, 85)
  ElseIf PocetOpTretiList = 2 Then
    Call Cara(345, 65, 480, 65)
    Call Sipka(345, 65, 345, 85)
    Call Sipka(480, 65, 480, 85)
    Call Cara(345, 300, 480, 300)
  ElseIf PocetOpTretiList = 3 Then
    Call Cara(278, 65, 548, 65)
    Call Sipka(413, 65, 413, 85)
    Call Sipka(278, 65, 278, 85)
    Call Sipka(548, 65, 548, 85)
    Call Cara(278, 300, 548, 300)
  ElseIf PocetOpTretiList = 4 Then
    Call Cara(210, 65, 615, 65)
    Call Sipka(210, 65, 210, 85)
    Call Sipka(345, 65, 345, 85)
    Call Sipka(480, 65, 480, 85)
    Call Sipka(615, 65, 615, 85)
    Call Cara(210, 300, 615, 300)
  ElseIf PocetOpTretiList = 5 Then
    Call Cara(143, 65, 683, 65)
    Call Sipka(413, 65, 413, 85)
    Call Sipka(143, 65, 143, 85)
    Call Sipka(278, 65, 278, 85)
    Call Sipka(548, 65, 548, 85)
    Call Sipka(683, 65, 683, 85)
    Call Cara(143, 300, 683, 300)
  End If
  
  For i = 1 To PocetOpTretiList
    If PocetOpTretiList = 1 Then
      Call TretiListOp(0, 1)
    ElseIf PocetOpTretiList = 2 Then
      If i = 1 Then Call TretiListOp(-68, 1)
      If i = 2 Then Call TretiListOp(68, 2)
    ElseIf PocetOpTretiList = 3 Then
      If i = 1 Then Call TretiListOp(-135, 1)
      If i = 2 Then Call TretiListOp(0, 2)
      If i = 3 Then Call TretiListOp(135, 3)
    ElseIf PocetOpTretiList = 4 Then
      If i = 1 Then Call TretiListOp(-203, 1)
      If i = 2 Then Call TretiListOp(-68, 2)
      If i = 3 Then Call TretiListOp(68, 3)
      If i = 4 Then Call TretiListOp(203, 4)
    ElseIf PocetOpTretiList = 5 Then
      If i = 1 Then Call TretiListOp(-270, 1)
      If i = 2 Then Call TretiListOp(-135, 2)
      If i = 3 Then Call TretiListOp(0, 3)
      If i = 4 Then Call TretiListOp(135, 4)
      If i = 5 Then Call TretiListOp(270, 5)
    End If
  Next i
  
  Call Sipka(413, 300, 413, 320)
  Call Rozhodnuti(313, 325, 200, 80)
  Call Textbox(340, 350, 150, 70, "Vyhovuje materiál specifikacím?" & vbNewLine & "Does the material match the specification?", 32, 33, 42, True, True)
  Call Sipka(518, 365, 557, 365)
  Call Textbox(510, 335, 50, 50, "Ne" & vbNewLine & "No", 2, 3, 3, True, False)
  Call Sipka(413, 410, 413, 430)
  Call Textbox(410, 403, 50, 50, "Ano" & vbNewLine & "Yes", 2, 3, 3, True, False)
  Call Dokument(562, 340, 100, 50)
  Call Textbox(562, 355, 100, 30, "OS 8.3-01", 9, 9, 0, False, False)
  Call Postup(363, 435, 100, 30)
  Call Textbox(363, 435, 100, 30, "Sklad V1" & vbNewLine & "Stock V1", 8, 9, 17, True, True)
  Call Dokument(363, 465, 100, 35)
  Call Textbox(363, 465, 100, 35, "Prùvodní lístek" & vbNewLine & "Carriage note", 0, 16, 28, True, False)
  Call Sipka(413, 505, 413, 530)
  Call Kruh(401, 535, 25, 25)
  Call Textbox(401, 539, 25, 25, "3", 1, 1, 1, False, False)
  
  ActiveSheet.Columns("A").EntireColumn.Delete
    
  With ActiveSheet.PageSetup
   .Zoom = False
   .FitToPagesTall = 1
   .FitToPagesWide = 1
  End With
End Sub

Sub TretiListOp(a As Double, c As Integer)
  Dim PPSvar As String
  Dim PPStocko As String
  Dim PPLumberg As String
  Dim PPTwist As String
  Dim d As Boolean
  Dim i As Integer
  
  If OpTretiList(c) = "Twistování vodièù" Then

    d = False
    PPTwist = Left(Dotaz.TxtCisloSvazku.Value, 7)
    For i = 1 To PocetPracovniPostupy
      If PracovniPostupy(i, 1) = "Twistování vodièù" Then
        If d = False Then
          PPTwist = PPTwist & Format(i, "00")
          d = True
        Else
          PPTwist = PPTwist & "/" & Format(i, "00")
        End If
      End If
    Next i

    If Posl = "Twistování vodièù" Then
      Call Postup(348 + a, 90, 130, 65)
      Call Dokument(348 + a, 155, 130, 140)
      Call Cara(413 + a, 295, 413 + a, 300)
      Call Textbox(348 + a, 240, 130, 10, "Balicí list", 0, 0, 0, False, False)
      Call Textbox(348 + a, 250, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
    Else
      Call Postup(348 + a, 90, 130, 65)
      Call Dokument(348 + a, 155, 130, 120)
      Call Cara(413 + a, 280, 413 + a, 300)
    End If
    Call TxtOperace(348 + a, 90, 130, 25, CislaOpTwist, Len(CislaOpTwist), 0, 0, False, True)
    Call Textbox(348 + a, 110, 130, 10, "Twistování vodièù", 17, 0, 0, False, True)
    Call Textbox(348 + a, 120, 130, 10, "Twist of wires", 14, 1, 14, True, True)
    Call Textbox(348 + a, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
    Call Textbox(348 + a, 140, 130, 10, "24-280", 0, 0, 0, False, False)
    Call Textbox(348 + a, 152, 130, 10, "Pracovní postup:", 0, 0, 0, False, False)
    Call Textbox(348 + a, 162, 130, 10, "Working plan:", 0, 1, 13, True, False)
    Call TxtBarText(348 + a, 173, 130, 10, PPTwist, 0, 0, 0, False, False)
    Call Textbox(348 + a, 183, 130, 20, "Kontrolní postup:" & vbNewLine & "Control plan:", 0, 18, 13, True, False)
    Call Textbox(348 + a, 204, 130, 30, Left(Dotaz.TxtCisloSvazku, 7) & "01", 0, 0, 0, False, False)
    Call Textbox(348 + a, 215, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
    
  ElseIf OpTretiList(c) = "Lisování kontaktu" Then
    Call Postup(348 + a, 90, 130, 65)
    Call Dokument(348 + a, 155, 130, 120)
    Call Cara(413 + a, 280, 413 + a, 300)
    If Posl = "Lisování kontaktu" Then
      Call Textbox(348 + a, 213, 130, 10, "Balicí list", 0, 0, 0, False, False)
      Call Textbox(348 + a, 223, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
    End If
    Call TxtOperace(348 + a, 90, 130, 25, CislaOpLisovani, Len(CislaOpLisovani), 0, 0, False, True)
    Call Textbox(348 + a, 110, 130, 10, "Lisování kontaktu", 17, 0, 0, False, True)
    Call Textbox(348 + a, 120, 130, 10, "Crimping of contact", 19, 1, 19, True, True)
    Call Textbox(348 + a, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
    Call Textbox(348 + a, 140, 130, 10, "Komax BT", 0, 0, 0, False, False)
    Call Textbox(348 + a, 155, 130, 90, "Typový kontrolní postup" & vbNewLine & "000.06" & vbNewLine & "Type control plan 000.06", 0, 31, 27, True, False)
    Call Textbox(348 + a, 190, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
    
  ElseIf OpTretiList(c) = "Svar vodièù" Then

    d = False
    PPSvar = Left(Dotaz.TxtCisloSvazku.Value, 7)
    For i = 1 To PocetPracovniPostupy
      If PracovniPostupy(i, 1) = "Svar vodièù" Then
        If d = False Then
          PPSvar = PPSvar & Format(i, "00")
          d = True
        Else
          PPSvar = PPSvar & "/" & Format(i, "00")
        End If
      End If
    Next i

    If Posl = "Svar vodièù" Then
      Call Postup(348 + a, 90, 130, 65)
      Call Dokument(348 + a, 155, 130, 140)
      Call Cara(413 + a, 295, 413 + a, 300)
      Call Textbox(348 + a, 243, 130, 10, "Balicí list", 0, 0, 0, False, False)
      Call Textbox(348 + a, 253, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
    Else
      Call Postup(348 + a, 90, 130, 65)
      Call Dokument(348 + a, 155, 130, 120)
      Call Cara(413 + a, 280, 413 + a, 300)
    End If
    Call TxtOperace(348 + a, 90, 130, 25, CislaOpSvar, Len(CislaOpSvar), 0, 0, False, True)
    Call Textbox(348 + a, 110, 130, 10, "Svar vodièù", 11, 0, 0, False, True)
    Call Textbox(348 + a, 120, 130, 10, "Ultrasonic welding", 18, 1, 18, True, True)
    Call Textbox(348 + a, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
    Call Textbox(348 + a, 140, 130, 10, "24-900", 0, 0, 0, False, False)
    Call Textbox(348 + a, 152, 130, 10, "Pracovní postup:", 0, 0, 0, False, False)
    Call Textbox(348 + a, 162, 130, 10, "Working plan:", 0, 1, 13, True, False)
    Call TxtBarText(348 + a, 173, 130, 10, PPSvar, 0, 0, 0, False, False)
    Call Textbox(348 + a, 185, 130, 90, "Typový kontrolní postup" & vbNewLine & "000.09" & vbNewLine & "Type control plan 000.09", 0, 31, 27, True, False)
    Call Textbox(348 + a, 220, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
    
  ElseIf OpTretiList(c) = "Osazování konektoru STOCKO" Then

    d = False
    PPStocko = Left(Dotaz.TxtCisloSvazku.Value, 7)
    For i = 1 To PocetPracovniPostupy
      If PracovniPostupy(i, 1) = "Osazování konektoru STOCKO" Then
        If d = False Then
          PPStocko = PPStocko & Format(i, "00")
          d = True
        Else
          PPStocko = PPStocko & "/" & Format(i, "00")
        End If
      End If
    Next i

    If Posl = "Osazování konektoru STOCKO" Then
      Call Postup(348 + a, 90, 130, 65)
      Call Dokument(348 + a, 155, 130, 140)
      Call Cara(413 + a, 295, 413 + a, 300)
      Call Textbox(348 + a, 243, 130, 10, "Balicí list", 0, 0, 0, False, False)
      Call Textbox(348 + a, 253, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
    Else
      Call Postup(348 + a, 90, 130, 65)
      Call Dokument(348 + a, 155, 130, 120)
      Call Cara(413 + a, 280, 413 + a, 300)
    End If
    Call TxtOperace(348 + a, 90, 130, 25, CislaOpStocko, Len(CislaOpStocko), 0, 0, False, True)
    Call Textbox(348 + a, 110, 130, 10, "Osazování konektoru STOCKO", 26, 0, 0, False, True)
    Call Textbox(348 + a, 120, 130, 10, "STOCKO connector`s assembly", 27, 1, 27, True, True)
    Call Textbox(348 + a, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
    Call Textbox(348 + a, 140, 130, 10, "30-003", 0, 0, 0, False, False)
    Call Textbox(348 + a, 152, 130, 10, "Pracovní postup:", 0, 0, 0, False, False)
    Call Textbox(348 + a, 162, 130, 10, "Working plan:", 0, 1, 13, True, False)
    Call TxtBarText(348 + a, 173, 130, 10, PPStocko, 0, 0, 0, False, False)
    Call Textbox(348 + a, 185, 130, 90, "Typový kontrolní postup" & vbNewLine & "000.10" & vbNewLine & "Type control plan 000.10", 0, 31, 27, True, False)
    Call Textbox(348 + a, 220, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
    
  ElseIf OpTretiList(c) = "Osazování konektoru LUMBERG" Then

    d = False
    PPLumberg = Left(Dotaz.TxtCisloSvazku.Value, 7)
    For i = 1 To PocetPracovniPostupy
      If PracovniPostupy(i, 1) = "Osazování konektoru LUMBERG" Then
        If d = False Then
          PPLumberg = PPLumberg & Format(i, "00")
          d = True
        Else
          PPLumberg = PPLumberg & "/" & Format(i, "00")
        End If
      End If
    Next i

    If Posl = "Osazování konektoru LUMBERG" Then
      Call Postup(348 + a, 90, 130, 65)
      Call Dokument(348 + a, 155, 130, 140)
      Call Cara(413 + a, 295, 413 + a, 300)
      Call Textbox(348 + a, 243, 130, 10, "Balicí list", 0, 0, 0, False, False)
      Call Textbox(348 + a, 253, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
    Else
      Call Postup(348 + a, 90, 130, 65)
      Call Dokument(348 + a, 155, 130, 120)
      Call Cara(413 + a, 280, 413 + a, 300)
    End If
    Call TxtOperace(348 + a, 90, 130, 25, CislaOpLumberg, Len(CislaOpLumberg), 0, 0, False, True)
    Call Textbox(348 + a, 110, 130, 10, "Osazování kon. LUMBERG", 22, 0, 0, False, True)
    Call Textbox(348 + a, 120, 130, 10, "LUMBERG connector`s assy", 24, 1, 20, True, True)
    Call Textbox(348 + a, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
    Call Textbox(348 + a, 140, 130, 10, "30-001", 0, 0, 0, False, False)
    Call Textbox(348 + a, 152, 130, 10, "Pracovní postup:", 0, 0, 0, False, False)
    Call Textbox(348 + a, 162, 130, 10, "Working plan:", 0, 1, 13, True, False)
    Call TxtBarText(348 + a, 173, 130, 10, PPLumberg, 0, 0, 0, False, False)
    Call Textbox(348 + a, 185, 130, 90, "Typový kontrolní postup" & vbNewLine & "000.11" & vbNewLine & "Type control plan 000.11", 0, 31, 27, True, False)
    Call Textbox(348 + a, 220, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)

  End If
End Sub

Sub Listy()
  Dim Name As String
  Dim list As Worksheet
  Set list = Worksheets.Add(After:=Worksheets(Worksheets.Count))
  list.Name = "VD" & Worksheets.Count - 1
  list.Select

  ActiveSheet.PageSetup.Orientation = xlPortrait
  
  Call ModObjekty.Kruh(201, 15, 25, 25)
  Call Textbox(201, 18, 25, 25, Mid(list.Name, 3, 2) - 1, 1, 1, 1, False, False)
  Call Dokument(148, 155, 130, 170)
  Call Sipka(213, 320, 213, 340)
  Call Rozhodnuti(113, 345, 200, 80)
  Call Textbox(140, 370, 150, 70, "Vyhovuje materiál specifikacím?" & vbNewLine & "Does the material match the specification?", 32, 33, 42, True, True)
  Call Sipka(318, 385, 357, 385)
  Call Textbox(310, 355, 50, 50, "Ne" & vbNewLine & "No", 2, 3, 3, True, False)
  Call Sipka(213, 430, 213, 450)
  Call Textbox(210, 423, 50, 50, "Ano" & vbNewLine & "Yes", 2, 3, 3, True, False)
  Call Dokument(362, 360, 100, 50)
  Call Textbox(362, 375, 100, 30, "OS 8.3-01", 9, 9, 0, False, False)
  Call Postup(163, 455, 100, 30)
  Call Textbox(163, 455, 100, 30, "Sklad V1" & vbNewLine & "Stock V1", 8, 9, 17, True, True)
  Call Dokument(163, 485, 100, 35)
  Call Textbox(163, 485, 100, 35, "Prùvodní lístek" & vbNewLine & "Carriage note", 0, 16, 28, True, False)
  Call Sipka(213, 525, 213, 550)
  Call Kruh(201, 555, 25, 25)
  Call Textbox(201, 559, 25, 25, Mid(list.Name, 3, 2), 1, 1, 1, False, False)
   
  With ActiveSheet.PageSetup
   .Zoom = False
   .FitToPagesTall = 1
   .FitToPagesWide = 1
  End With
 
End Sub

Sub ZapisStul()
  Dim i As Integer
  Dim d As Boolean
  Dim PPstul As String
    
  d = False
  PPstul = Left(Dotaz.TxtCisloSvazku.Value, 7)
  For i = 1 To PocetPracovniPostupy
    If PracovniPostupy(i, 1) = "Montáž podskupiny" Then
      If d = False Then
        PPstul = PPstul & Format(i, "00")
        d = True
      Else
        PPstul = PPstul & "/" & Format(i, "00")
      End If
    End If
  Next i
  
  Call ModNacteniOperaci.ZapisOperaci(PocetStul, "Montáž podskupiny", CislaOpStul, Stul)
  Call Listy
  Call ModObjekty.Sipka(213, 45, 213, 85)
  Call Postup(148, 90, 130, 65)
  Call TxtOperace(148, 90, 130, 25, CislaOpStul, Len(CislaOpStul), 0, 0, False, True)
  Call Textbox(148, 110, 130, 10, "Montáž podskupiny", 17, 0, 0, False, True)
  Call Textbox(148, 120, 130, 10, "Assembly of subset", 18, 1, 18, True, True)
  Call Textbox(148, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
  Call Textbox(148, 140, 130, 10, "24-500", 0, 0, 0, False, False)
  Call Textbox(148, 155, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
  Call Textbox(148, 178, 130, 10, "Pracovní postup:", 0, 0, 0, False, False)
  Call Textbox(148, 188, 130, 10, "Working plan:", 0, 1, 13, True, False)
  Call TxtBarText(148, 198, 130, 10, PPstul, 0, 0, 0, False, False)
  Call Textbox(148, 210, 130, 20, "Kontrolní postup:", 0, 0, 0, True, False)
  Call Textbox(148, 220, 130, 20, "Control plan:", 0, 1, 13, True, False)
  Call Textbox(148, 230, 130, 30, Left(Dotaz.TxtCisloSvazku, 7) & "01", 0, 0, 0, False, False)
  If Posl = "Montáž podskupiny" Then
    Call Textbox(148, 242, 130, 10, "Balicí list", 0, 0, 0, False, False)
    Call Textbox(148, 252, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
  End If

  ActiveSheet.Columns("A").EntireColumn.Delete
  
End Sub

Sub ZapisRBK()
  Dim i As Integer
  Dim PPRBK As String
  Dim d As Boolean
     
  d = False
  PPRBK = Left(Dotaz.TxtCisloSvazku.Value, 7)
  For i = 1 To PocetPracovniPostupy
    If PracovniPostupy(i, 1) = "Zafoukání ITRs" Then
      If d = False Then
        PPRBK = PPRBK & Format(i, "00")
        d = True
      Else
        PPRBK = PPRBK & "/" & Format(i, "00")
      End If
    End If
  Next i

  Call ModNacteniOperaci.ZapisOperaci(PocetRBK, "Zafoukání ITRs", CislaOpRBK, RBK)
  Call Listy
  Call ModObjekty.Sipka(213, 45, 213, 85)
  Call Postup(148, 90, 130, 65)
  Call TxtOperace(148, 90, 130, 25, CislaOpRBK, Len(CislaOpRBK), 0, 0, False, True)
  Call Textbox(148, 110, 130, 10, "Zafoukání ITRs", 14, 0, 0, False, True)
  Call Textbox(148, 120, 130, 10, "Shrinking of insulation tube", 27, 1, 27, True, True)
  Call Textbox(148, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
  Call Textbox(148, 140, 130, 10, "24-530", 0, 0, 0, False, False)
  Call Textbox(148, 155, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
  Call Textbox(148, 178, 130, 10, "Pracovní postup:", 0, 0, 0, False, False)
  Call Textbox(148, 188, 130, 10, "Working plan:", 0, 1, 13, True, False)
  Call TxtBarText(148, 198, 130, 10, PPRBK, 0, 0, 0, False, False)
  Call Textbox(148, 210, 130, 20, "Kontrolní postup:", 0, 0, 0, True, False)
  Call Textbox(148, 220, 130, 20, "Control plan:", 0, 1, 13, True, False)
  Call Textbox(148, 230, 130, 30, Left(Dotaz.TxtCisloSvazku, 7) & "01", 0, 0, 0, False, False)
  If Posl = "Zafoukání ITRs" Then
    Call Textbox(148, 242, 130, 10, "Balicí list", 0, 0, 0, False, False)
    Call Textbox(148, 252, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
  End If
  
  ActiveSheet.Columns("A").EntireColumn.Delete
  
End Sub

Sub ZapisBandaz()
  Dim i As Integer
  Dim PPBandaz As String
  Dim KaPrBandaz As String
  Dim Pracoviste As String
  Dim d As Boolean
  Dim index As Integer
     
  d = False
  PPBandaz = Left(Dotaz.TxtCisloSvazku.Value, 7)
  For i = 1 To PocetPracovniPostupy
    If PracovniPostupy(i, 1) = "Strojní bandáž" Then
      If d = False Then
        PPBandaz = PPBandaz & Format(i, "00")
        d = True
      Else
        PPBandaz = PPBandaz & "/" & Format(i, "00")
      End If
    End If
  Next i
      
  d = False
  KaPrBandaz = Left(Dotaz.TxtCisloSvazku.Value, 7)
  For i = 1 To PocetKaPr
    If KartyProcesu(i) = "Strojní bandáž" Then
      If d = False Then
        KaPrBandaz = KaPrBandaz & Format(i, "00")
        d = True
      Else
        KaPrBandaz = KaPrBandaz & "/" & Format(i, "00")
      End If
    End If
  Next i
      
  For index = 1 To PocetBandaz
    If index = 1 Then
      Pracoviste = OpBandaz(index)
    ElseIf index > 1 Then
      Pracoviste = Pracoviste + ", " + OpBandaz(index)
    End If
  Next index

  Call ModNacteniOperaci.ZapisOperaci(PocetBandaz, "Strojní bandáž", CislaOpBandaz, Bandaz)
  Call Listy
  Call ModObjekty.Sipka(213, 45, 213, 85)
  Call Postup(148, 90, 130, 65)
  Call TxtOperace(148, 90, 130, 25, CislaOpBandaz, Len(CislaOpBandaz), 0, 0, False, True)
  Call TxtBarText(148, 110, 130, 10, "Strojní bandáž", 14, 0, 0, False, True)
  Call TxtBarText(148, 120, 130, 10, "Automatic taping", 16, 1, 16, True, True)
  Call Textbox(148, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
  Call Textbox(148, 140, 130, 10, Pracoviste, 0, 0, 0, False, False)
  Call Textbox(148, 155, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
  Call Textbox(148, 178, 130, 10, "Pracovní postup:", 0, 0, 0, False, False)
  Call Textbox(148, 188, 130, 10, "Working plan:", 0, 1, 13, True, False)
  Call TxtBarText(148, 198, 130, 10, PPBandaz, 0, 0, 0, False, False)
  Call Textbox(148, 210, 130, 20, "Kontrolní postup:", 0, 0, 0, True, False)
  Call Textbox(148, 220, 130, 20, "Control plan:", 0, 1, 13, True, False)
  Call Textbox(148, 230, 130, 30, Left(Dotaz.TxtCisloSvazku, 7) & "01", 0, 0, 0, False, False)
  Call Textbox(148, 242, 130, 10, "Karta procesu:", 0, 0, 0, False, False)
  Call Textbox(148, 252, 130, 10, "Process card:", 0, 1, 13, True, False)
  Call TxtBarText(148, 262, 130, 10, KaPrBandaz, 0, 0, 0, False, False)
  If Posl = "Strojní bandáž" Then
    Call Textbox(148, 274, 130, 10, "Balicí list", 0, 0, 0, False, False)
    Call Textbox(148, 284, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
  End If
  
  ActiveSheet.Columns("A").EntireColumn.Delete
  
End Sub

Sub ZapisZastrik()
  Dim i As Integer
  Dim PPZastrik As String
  Dim Pracoviste As String
  Dim d As Boolean
  Dim index As Integer
      
  d = False
  PPZastrik = Left(Dotaz.TxtCisloSvazku.Value, 7)
  For i = 1 To PocetPracovniPostupy
    If PracovniPostupy(i, 1) = "Nízkotlaký zástøik" Then
      If d = False Then
        PPZastrik = PPZastrik & Format(i, "00")
        d = True
      Else
        PPZastrik = PPZastrik & "/" & Format(i, "00")
      End If
    End If
  Next i
    
  For index = 1 To PocetZastrik
    If index = 1 Then
      Pracoviste = OpZastrik(index)
    ElseIf index > 1 Then
      Pracoviste = Pracoviste + ", " + OpZastrik(index)
    End If
  Next index

  Call ModNacteniOperaci.ZapisOperaci(PocetZastrik, "Nízkotlaký zástøik", CislaOpZastrik, Zastrik)
  Call Listy
  Call ModObjekty.Sipka(213, 45, 213, 85)
  Call Postup(148, 90, 130, 65)
  Call TxtOperace(148, 90, 130, 25, CislaOpZastrik, Len(CislaOpZastrik), 0, 0, False, True)
  Call Textbox(148, 110, 130, 10, "Nízkotlaký zástøik", 18, 0, 0, False, True)
  Call Textbox(148, 120, 130, 10, "Low-pressure molding", 20, 1, 20, True, True)
  Call Textbox(148, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
  Call Textbox(148, 140, 130, 10, Pracoviste, 0, 0, 0, False, False)
  Call Textbox(148, 155, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
  Call Textbox(148, 178, 130, 10, "Pracovní postup:", 0, 0, 0, False, False)
  Call Textbox(148, 188, 130, 10, "Working plan:", 0, 1, 13, True, False)
  Call TxtBarText(148, 198, 130, 10, PPZastrik, 0, 0, 0, False, False)
  Call Textbox(148, 210, 130, 20, "Kontrolní postup:", 0, 0, 0, True, False)
  Call Textbox(148, 220, 130, 20, "Control plan:", 0, 1, 13, True, False)
  Call Textbox(148, 230, 130, 30, Left(Dotaz.TxtCisloSvazku, 7) & "01", 0, 0, 0, False, False)
  Call Textbox(148, 242, 130, 10, "Elektronická karta procesu", 0, 0, 0, False, False)
  Call Textbox(148, 252, 130, 10, "Elektronic process card", 0, 1, 23, True, False)
  If Posl = "Nízkotlaký zástøik" Then
    Call Textbox(148, 274, 130, 10, "Balicí list", 0, 0, 0, False, False)
    Call Textbox(148, 284, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
  End If

  ActiveSheet.Columns("A").EntireColumn.Delete

End Sub

Sub ZapisRAMPF()
  Dim i As Integer
  Dim PPRAMPF As String
  Dim Pracoviste As String
  Dim d As Boolean
  Dim index As Integer

  d = False
  PPRAMPF = Left(Dotaz.TxtCisloSvazku.Value, 7)
  For i = 1 To PocetPracovniPostupy
    If PracovniPostupy(i, 1) = "RAMPF" Then
      If d = False Then
        PPRAMPF = PPRAMPF & Format(i, "00")
        d = True
      Else
        PPRAMPF = PPRAMPF & "/" & Format(i, "00")
      End If
    End If
  Next i

  Call ModNacteniOperaci.ZapisOperaci(PocetRAMPF, "RAMPF", CislaOpRAMPF, RAMPF)
  Call Listy
  Call ModObjekty.Sipka(213, 45, 213, 85)
  Call Postup(148, 90, 130, 65)
  Call TxtOperace(148, 90, 130, 25, CislaOpRAMPF, Len(CislaOpRAMPF), 0, 0, False, True)
  Call Textbox(148, 110, 130, 10, "Zalévání", 8, 0, 0, False, True)
  Call Textbox(148, 120, 130, 10, "Potting", 7, 1, 18, True, True)
  Call Textbox(148, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
  Call Textbox(148, 140, 130, 10, "38-001", 0, 0, 0, False, False)
  Call Textbox(148, 155, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
  Call Textbox(148, 178, 130, 10, "Pracovní postup:", 0, 0, 0, False, False)
  Call Textbox(148, 188, 130, 10, "Working plan:", 0, 1, 13, True, False)
  Call TxtBarText(148, 198, 130, 10, PPRAMPF, 0, 0, 0, False, False)
  Call Textbox(148, 210, 130, 20, "Kontrolní postup:", 0, 0, 0, True, False)
  Call Textbox(148, 220, 130, 20, "Control plan:", 0, 1, 13, True, False)
  Call Textbox(148, 230, 130, 30, Left(Dotaz.TxtCisloSvazku, 7) & "01", 0, 0, 0, False, False)
  Call Textbox(148, 242, 130, 10, "Elektronická karta procesu", 0, 0, 0, False, False)
  Call Textbox(148, 252, 130, 10, "Elektronic process card", 0, 1, 23, True, False)
  If Posl = "RAMPF" Then
      Call Textbox(148, 274, 130, 10, "Balicí list", 0, 0, 0, False, False)
      Call Textbox(148, 284, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
  End If
  
  ActiveSheet.Columns("A").EntireColumn.Delete

End Sub

Sub ZapisMontaz(i As Integer)
  
  Dim e As Integer
  Dim PPMontaz As String
  Dim KaPrMontaz As String
  Dim d As Boolean

  PPMontaz = Left(Dotaz.TxtCisloSvazku.Value, 7) & Format(i, "00")
  
  For e = 1 To PocetKaPr
    If KartyProcesu(e) = PracovniPostupy(i, 4) Then KaPrMontaz = Left(Dotaz.TxtCisloSvazku.Value, 7) & Format(e, "00")
  Next e
  
  Call Listy

  If Len(CStr(PracovniPostupy(i, 4))) Then
    Call ModObjekty.Sipka(213, 45, 213, 55)
    Call Postup(148, 60, 130, 95)
    Call TxtOperace(148, 60, 130, 25, "OP. " & CStr(PracovniPostupy(i, 2)), Len(PracovniPostupy(i, 2)) + 4, 0, 0, False, True)
    Call Textbox(148, 80, 130, 10, CStr(PracovniPostupy(i, 4)), Len(PracovniPostupy(i, 4)), 0, 0, False, True)
    If Preklad <> "" Then Call TxtBarText(148, 105, 130, 10, Preklad, Len(Preklad), 1, Len(Preklad), True, True)
  Else
    Call ModObjekty.Sipka(213, 45, 213, 85)
    Call Postup(148, 90, 130, 65)
    Call TxtOperace(148, 90, 130, 25, "OP. " & CStr(PracovniPostupy(i, 2)), Len(PracovniPostupy(i, 2)) + 4, 0, 0, False, True)
    Call Textbox(148, 110, 130, 10, CStr(PracovniPostupy(i, 4)), Len(PracovniPostupy(i, 4)), 0, 0, False, True)
    If Preklad <> "" Then Call TxtBarText(148, 120, 130, 10, Preklad, Len(Preklad), 1, Len(Preklad), True, True)
  End If
  Call Textbox(148, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
  Call Textbox(148, 140, 130, 10, CStr(PracovniPostupy(i, 3)), 0, 0, 0, False, False)
  Call Textbox(148, 155, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
  Call Textbox(148, 178, 130, 10, "Pracovní postup:", 0, 0, 0, False, False)
  Call Textbox(148, 188, 130, 10, "Working plan:", 0, 1, 13, True, False)
  Call TxtBarText(148, 198, 130, 10, PPMontaz, 0, 0, 0, False, False)
  Call Textbox(148, 210, 130, 20, "Kontrolní postup:", 0, 0, 0, True, False)
  Call Textbox(148, 220, 130, 20, "Control plan:", 0, 1, 13, True, False)
  Call Textbox(148, 230, 130, 30, Left(Dotaz.TxtCisloSvazku, 7) & "01", 0, 0, 0, False, False)
  Call Textbox(148, 242, 130, 10, "Karta procesu:", 0, 0, 0, False, False)
  Call Textbox(148, 252, 130, 10, "Process card:", 0, 1, 13, True, False)
  Call TxtBarText(148, 262, 130, 10, KaPrMontaz, 0, 0, 0, False, False)
  If Posl = PracovniPostupy(i, 4) Then
    Call Textbox(148, 274, 130, 10, "Balicí list", 0, 0, 0, False, False)
    Call Textbox(148, 284, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
  End If
        
  ActiveSheet.Columns("A").EntireColumn.Delete
      
End Sub

Sub ZapisLaser()
  Dim PPLaser As String
  Dim KaPrLaser As String
  Dim d As Boolean
  Dim i As Integer
    
  d = False
  PPLaser = Left(Dotaz.TxtCisloSvazku.Value, 7)
  For i = 1 To PocetPracovniPostupy
    If PracovniPostupy(i) = "Laserové znaèení" Then
      If d = False Then
        PPLaser = PPLaser & Format(i, "00")
        d = True
      Else
        PPLaser = PPLaser & "/" & Format(i, "00")
      End If
    End If
  Next i
  
  d = False
  KaPrLaser = Left(Dotaz.TxtCisloSvazku.Value, 7)
  For i = 1 To PocetKaPr
    If KartyProcesu(i) = "Laserové znaèení" Then
      If d = False Then
        KaPrLaser = KaPrLaser & Format(i, "00")
        d = True
      Else
        KaPrLaser = KaPrLaser & "/" & Format(i, "00")
      End If
    End If
  Next i

  Call ModNacteniOperaci.ZapisOperaci(PocetLaser, "Laserové znaèení", CislaOpLaser, Laser)
  Call Listy
  Call ModObjekty.Sipka(213, 45, 213, 85)
  Call Postup(148, 90, 130, 65)
  Call TxtOperace(148, 90, 130, 25, CislaOpLaser, Len(CislaOpLaser), 0, 0, False, True)
  Call Textbox(148, 110, 130, 10, "Laserové znaèení", 16, 0, 0, False, True)
  Call Textbox(148, 120, 130, 10, "Laser marking", 13, 1, 13, True, True)
  Call Textbox(148, 130, 130, 10, "Pracovištì/Work place:", 0, 11, 12, True, False)
  Call Textbox(148, 140, 130, 10, "40-001", 0, 0, 0, False, False)
  Call Textbox(148, 155, 130, 30, "Mzdový lístek" & vbNewLine & "Operation card", 0, 14, 15, True, False)
  Call Textbox(148, 178, 130, 10, "Pracovní postup:", 0, 0, 0, False, False)
  Call Textbox(148, 188, 130, 10, "Working plan:", 0, 1, 13, True, False)
  Call TxtBarText(148, 198, 130, 10, PPLaser, 0, 0, 0, False, False)
  Call Textbox(148, 210, 130, 20, "Kontrolní postup:", 0, 0, 0, True, False)
  Call Textbox(148, 220, 130, 20, "Control plan:", 0, 1, 13, True, False)
  Call Textbox(148, 230, 130, 30, Left(Dotaz.TxtCisloSvazku, 7) & "01", 0, 0, 0, False, False)
  Call Textbox(148, 242, 130, 10, "Karta procesu:", 0, 0, 0, False, False)
  Call Textbox(148, 252, 130, 10, "Process card:", 0, 1, 13, True, False)
  Call TxtBarText(148, 262, 130, 10, KaPrLaser, 0, 0, 0, False, False)
  If Posl = "Laserové znaèení" Then
    Call Textbox(148, 274, 130, 10, "Balicí list", 0, 0, 0, False, False)
    Call Textbox(148, 284, 130, 10, "Packaging sheet", 0, 1, 15, True, False)
  End If
  
  ActiveSheet.Columns("A").EntireColumn.Delete

End Sub

Sub PosledniList()
  Dim Name As String
  Dim list As Worksheet
  Set list = Worksheets.Add(After:=Worksheets(Worksheets.Count))
  list.Name = "VD" & Worksheets.Count - 1
  list.Select

  ActiveSheet.PageSetup.Orientation = xlPortrait

  Call ModObjekty.Kruh(201, 15, 25, 25)
  Call Textbox(201, 18, 25, 25, Mid(list.Name, 3, 2) - 1, 1, 1, 1, False, False)
  Call Sipka(213, 45, 213, 65)
  Call Postup(148, 70, 130, 40)
  Call Textbox(148, 75, 130, 40, "Výstupní kontrola" & vbNewLine & "Outgoing control", 34, 18, 17, True, True)
  Call Dokument(148, 110, 130, 160)
  Call Textbox(148, 110, 130, 30, "Prùvodní lístek" & vbNewLine & "Carriage note", 0, 16, 13, True, False)
  Call Textbox(148, 135, 130, 30, "Výkres výrobku" & vbNewLine & "Product drawing", 0, 14, 15, True, False)
  Call Textbox(148, 160, 130, 10, "Typový kontrolní postup:" & vbNewLine & "Type control plan:" & vbNewLine & "001.99", 0, 25, 18, True, False)
  Call Textbox(148, 195, 130, 10, "Balicí list" & vbNewLine & "Packaging sheet" & vbNewLine & "OS 8.1-01", 0, 12, 15, True, False)
  Call Sipka(213, 270, 213, 295)
  Call Rozhodnuti(113, 300, 200, 80)
  Call Textbox(140, 325, 150, 70, "Vyhovuje materiál specifikacím?" & vbNewLine & "Does the material match the specification?", 32, 33, 42, True, True)
  Call Sipka(318, 340, 357, 340)
  Call Textbox(310, 310, 50, 50, "Ne" & vbNewLine & "No", 2, 3, 3, True, False)
  Call Sipka(213, 385, 213, 405)
  Call Textbox(210, 378, 50, 50, "Ano" & vbNewLine & "Yes", 2, 3, 3, True, False)
  Call Textbox(60, 378, 150, 50, "Paletový vozík/Pallet truck", 0, 16, 12, True, False)
  Call Textbox(60, 388, 150, 50, "Vysokozvižný vozík/Fork lifter", 0, 20, 11, True, False)
  Call Dokument(362, 315, 100, 50)
  Call Textbox(362, 330, 100, 30, "OS 8.3-01", 9, 9, 0, False, False)
  Call Postup(148, 410, 130, 30)
  Call Textbox(148, 410, 130, 30, "Expedièní sklad EX" & vbNewLine & "Dispatch stock EX", 18, 19, 18, True, True)
  Call Dokument(148, 440, 130, 70)
  Call Textbox(148, 440, 130, 35, "Prùvodní lístek" & vbNewLine & "Carriage note", 0, 16, 28, True, False)
  Call Textbox(148, 465, 130, 35, "Identifikaèní lístek - Pøíjem" & vbNewLine & "Identification sheet - Receipt", 0, 30, 30, True, False)
  Call Sipka(213, 510, 213, 540)
  Call Textbox(60, 512, 150, 50, "Paletový vozík/Pallet truck", 0, 16, 12, True, False)
  Call Textbox(60, 522, 150, 50, "Vysokozvižný vozík/Fork lifter", 0, 20, 11, True, False)
  Call Postup(148, 545, 130, 30)
  Call Textbox(148, 545, 130, 30, "Expedice" & vbNewLine & "Export", 8, 9, 6, True, True)
  Call Dokument(148, 575, 130, 130)
  Call Textbox(148, 575, 130, 35, "Prùvodní lístek" & vbNewLine & "Carriage note", 0, 16, 28, True, False)
  Call Textbox(148, 600, 130, 35, "Identifikaèní lístek - Pøíjem" & vbNewLine & "Identification sheet - Receipt", 0, 30, 30, True, False)
  Call Textbox(148, 625, 130, 35, "Balicí list" & vbNewLine & "Packaging sheet", 0, 12, 15, True, False)
  Call Textbox(148, 650, 130, 35, "Dodací list" & vbNewLine & "Delivery note", 0, 12, 13, True, False)
  
  ActiveSheet.Columns("A").EntireColumn.Delete
  
  With ActiveSheet.PageSetup
   .Zoom = False
   .FitToPagesTall = 1
   .FitToPagesWide = 1
  End With
End Sub
                                                                                                                                                                                                                                                                         