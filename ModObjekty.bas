Attribute VB_Name = "ModObjekty"
Option Explicit

Sub Kruh(a As Single, b As Single, c As Single, d As Single)
  With ActiveSheet.Shapes.AddShape(msoShapeOval, a, b, c, d)
    .Fill.ForeColor.RGB = RGB(230, 230, 230)
    .Line.ForeColor.RGB = RGB(0, 0, 0)
    .Line.Weight = 1
  End With
End Sub

Sub Postup(a As Single, b As Single, c As Single, d As Single)
  With ActiveSheet.Shapes.AddShape(msoShapeFlowchartProcess, a, b, c, d)
    .Fill.ForeColor.RGB = RGB(230, 230, 230)
    .Line.ForeColor.RGB = RGB(0, 0, 0)
    .Line.Weight = 1
  End With
End Sub

Sub Cara(a As Single, b As Single, c As Single, d As Single)
  With ActiveSheet.Shapes.AddConnector(msoConnectorStraight, a, b, c, d)
    .Line.Weight = 1
    .Line.ForeColor.RGB = RGB(0, 0, 0)
    .Line.Weight = 1
  End With
End Sub

Sub Sipka(a As Single, b As Single, c As Single, d As Single)
  With ActiveSheet.Shapes.AddConnector(msoConnectorStraight, a, b, c, d)
    .Line.EndArrowheadStyle = msoArrowheadTriangle
    .Line.Weight = 1
    .Line.ForeColor.RGB = RGB(0, 0, 0)
    .Line.Weight = 1
  End With
End Sub

Sub Dokument(a As Single, b As Single, c As Single, d As Single)
  With ActiveSheet.Shapes.AddShape(msoShapeFlowchartDocument, a, b, c, d)
    .Fill.ForeColor.RGB = RGB(230, 230, 230)
    .Line.ForeColor.RGB = RGB(0, 0, 0)
    .Line.Weight = 1
  End With
End Sub

Sub Rozhodnuti(a As Single, b As Single, c As Single, d As Single)
  With ActiveSheet.Shapes.AddShape(msoShapeFlowchartDecision, a, b, c, d)
    .Fill.ForeColor.RGB = RGB(230, 230, 230)
    .Line.ForeColor.RGB = RGB(0, 0, 0)
    .Line.Weight = 1
  End With
End Sub

Sub Textbox(a As Single, b As Single, c As Single, d As Single, e As String, delkatucne As Single, zacatekkurzivy As Single, delkakurzivy As Single, kurziva As Boolean, tucne As Boolean)
  With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, a, b, c, d)
    .TextFrame.Characters.Text = e
    .TextFrame.Characters.Font.Size = 9
    .TextFrame.HorizontalAlignment = xlCenter
    .TextFrame.VerticalAlignment = xlCenter
    .TextFrame.Characters(Start:=0, Length:=delkatucne).Font.Bold = tucne
    .TextFrame.Characters(Start:=zacatekkurzivy, Length:=delkakurzivy).Font.Italic = kurziva
  End With
End Sub

Sub TxtOperace(a As Single, b As Single, c As Single, d As Single, e As String, delkatucne As Single, zacatekkurzivy As Single, delkakurzivy As Single, kurziva As Boolean, tucne As Boolean)
  If Len(e) < 26 Then b = b + 6
  With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, a, b, c, d)
    .TextFrame.Characters.Text = e
      If .TextFrame.Characters.Count > 50 Then
        .TextFrame.Characters.Font.Size = 7
      Else
        .TextFrame.Characters.Font.Size = 9
      End If
    .TextFrame.HorizontalAlignment = xlCenter
    .TextFrame.VerticalAlignment = xlCenter
    .TextFrame.Characters(Start:=0, Length:=delkatucne).Font.Bold = tucne
    .TextFrame.Characters(Start:=zacatekkurzivy, Length:=delkakurzivy).Font.Italic = kurziva
  End With
End Sub

Sub TxtBarText(a As Single, b As Single, c As Single, d As Single, e As String, delkatucne As Single, zacatekkurzivy As Single, delkakurzivy As Single, kurziva As Boolean, tucne As Boolean)
  With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, a, b, c, d)
    .TextFrame.Characters.Text = e
    .TextFrame.Characters.Font.Size = 9
    .TextFrame.Characters.Font.ColorIndex = 3
    .TextFrame.HorizontalAlignment = xlCenter
    .TextFrame.VerticalAlignment = xlCenter
    .TextFrame.Characters(Start:=0, Length:=delkatucne).Font.Bold = tucne
    .TextFrame.Characters(Start:=zacatekkurzivy, Length:=delkakurzivy).Font.Italic = kurziva
  End With
End Sub

Sub TxtPreklad(a As Single, b As Single, c As Single, d As Single, e As String, delkatucne As Single, zacatekkurzivy As Single, delkakurzivy As Single, kurziva As Boolean, tucne As Boolean)
  With ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, a, b, c, d)
    .TextFrame.Characters.Text = e
      If .TextFrame.Characters.Count > 24 Then
        .TextFrame.Characters.Font.Size = 7
      Else
        .TextFrame.Characters.Font.Size = 9
      End If
    .TextFrame.Characters.Font.ColorIndex = 3
    .TextFrame.HorizontalAlignment = xlCenter
    .TextFrame.VerticalAlignment = xlCenter
    .TextFrame.Characters(Start:=0, Length:=delkatucne).Font.Bold = tucne
    .TextFrame.Characters(Start:=zacatekkurzivy, Length:=delkakurzivy).Font.Italic = kurziva
  End With
End Sub
N T . D O M   U S E R D O M A I N _ R O A M I N G P R O F 