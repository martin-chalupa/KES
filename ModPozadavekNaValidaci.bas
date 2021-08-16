Attribute VB_Name = "ModPozadavekNaValidaci"
Option Explicit

Public CisloSvazku As String
Public CislaRadku As New Collection
Public PocetRadku As Long
Public PocetBunek As Integer
Public Bunka As Range
Public PrvniBunka As Long
Public ValidaceKont As New Collection
Public ValidaceVod As New Collection
Public ValidaceTes As New Collection
Public ValidaceDvojKont As New Collection
Public ValidaceDvojVod As New Collection
Public ValidaceX As Boolean
Public ValidaceY As Boolean
Public Pozadavek As Workbook
Public Norma As String
Public PocetValidaci As Integer

Sub VytvorPozadavek()
  FormPozadavek.Show
End Sub

Sub PozadavekNaValidaci()
Dim i As Long
Dim n As Integer
Dim a As Integer
Dim b As Integer

PocetRadku = ThisWorkbook.Sheets("DATA1").Cells(Rows.Count, 1).End(xlUp).Row

CisloSvazku = FormPozadavek.TxtCisloSvazku.Text

Set Bunka = ThisWorkbook.Sheets("DATA1").Range("A2:A" & PocetRadku).Find(What:=CisloSvazku)
If Bunka Is Nothing Then
  MsgBox "Svazek nebyl nalezen!"
  Exit Sub
End If

i = 1
If Not Bunka Is Nothing Then
  PrvniBunka = Bunka.Row
  Do
    PocetBunek = PocetBunek + 1
    CislaRadku.Add Bunka.Row
    i = i + 1
    Set Bunka = ThisWorkbook.Sheets("DATA1").Range("A2:A" & PocetRadku).FindNext(Bunka)
  Loop While Not Bunka Is Nothing And Bunka.Row <> PrvniBunka
End If

With ThisWorkbook.Sheets("DATA1")
  For i = 1 To PocetBunek
    If .Range("AO" & CislaRadku(i)).Value = "Chybí validaceX" Then
      ValidaceX = True
      Norma = .Range("F" & CislaRadku(i)).Value
      ValidaceKont.Add .Range("T" & CislaRadku.Item(i)).Value
      ValidaceVod.Add .Range("Z" & CislaRadku.Item(i)).Value
      If InStr(.Range("V" & CislaRadku.Item(i)).Value, "PARP") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "EHR") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "ARVPB12C122AK") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "70107-0045") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-1418968-1") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "XARP") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "14650.669.696") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "6098-5283") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "V0072020B111") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "PMS") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-1670988-1") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-0953119-1") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "XMS") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "KG928001") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-0953119-2") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-0929178-1") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-1534113-1") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-02208746-1") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-1394396-1") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "5Q0.973.115") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-02208748-1") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "716.06.301.00") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-0965490-2") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-2295410-1") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-1743282-1") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "2208960-1") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-2295408-1") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "756.01.104.30") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "ARVPB-18-3AK") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "502351") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "101702780001") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "ARVPB-12-2AK") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-2208746-1") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "101802780001") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "6098-5269") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "505151") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "42113100") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "6249-1243") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-2295163") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "42112700") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "6188-0555") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "2-0968326-1") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "161571-00") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "6188-5542") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "1-2292506-1") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "162183-00") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "6188-5544") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-1743282-1") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "1-1452337-1") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "100802780001") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "2208963-1") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "ARVWSB-18-3AK") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "100702780001") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "1-1452337-1") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "ARVWSB-12B-2AK") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "60013570A01C") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "1-2292507-1") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "1304487040") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "805-587-551") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "9-2208748-1") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "717.09.301.00") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "7283-5927-10") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "V001691012211") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "1304487044") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "33482-2140") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "28N010") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "42075100") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "K7920-9208") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "805-587-541") > 0 Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "42063900") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0007304") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "33482-2040") Or _
      InStr(.Range("V" & CislaRadku.Item(i)).Value, "ARVPB18C183AK") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-1456987-3") > 0 Or InStr(.Range("V" & CislaRadku.Item(i)).Value, "0-1418967-1") Then
        ValidaceTes.Add ""
      Else
        ValidaceTes.Add .Range("V" & CislaRadku.Item(i)).Value
      End If
    ElseIf .Range("AS" & CislaRadku(i)).Value = "Chybí validace X" Then
      ValidaceX = True
      Norma = .Range("F" & CislaRadku(i)).Value
      ValidaceDvojKont.Add .Range("T" & CislaRadku.Item(i)).Value
      If .Range("T" & CislaRadku.Item(i + 1)).Value = "" Then
        ValidaceDvojKont.Add "-"
      Else
        ValidaceDvojKont.Add .Range("T" & CislaRadku.Item(i + 1)).Value
      End If
      ValidaceDvojVod.Add .Range("Z" & CislaRadku.Item(i)).Value
      ValidaceDvojVod.Add .Range("Z" & CislaRadku.Item(i + 1)).Value
    End If
  Next i

  For i = 1 To PocetBunek
    If .Range("AP" & CislaRadku(i)).Value = "Chybí validaceY" Then
      ValidaceY = True
      Norma = .Range("F" & CislaRadku(i)).Value
      ValidaceKont.Add .Range("AH" & CislaRadku.Item(i)).Value
      ValidaceVod.Add .Range("Z" & CislaRadku.Item(i)).Value
      If InStr(.Range("AJ" & CislaRadku.Item(i)).Value, "PARP") > 0 Or InStr(.Range("AJ" & CislaRadku.Item(i)).Value, "EHR") > 0 Or InStr(.Range("AJ" & CislaRadku.Item(i)).Value, "XARP") > 0 Or InStr(.Range("AJ" & CislaRadku.Item(i)).Value, "PMS") > 0 Or InStr(.Range("AJ" & CislaRadku.Item(i)).Value, "XMS") > 0 Then
        ValidaceTes.Add ""
      Else
        ValidaceTes.Add .Range("AJ" & CislaRadku.Item(i)).Value
      End If
    ElseIf .Range("AT" & CislaRadku(i)).Value = "Chybí validace Y" Then
      ValidaceY = True
      Norma = .Range("F" & CislaRadku(i)).Value
      ValidaceDvojKont.Add .Range("AH" & CislaRadku.Item(i)).Value
      If .Range("AH" & CislaRadku.Item(i + 1)).Value = "" Then
        ValidaceDvojKont.Add "-"
      Else
        ValidaceDvojKont.Add .Range("AH" & CislaRadku.Item(i + 1)).Value
      End If
      ValidaceDvojVod.Add .Range("Z" & CislaRadku.Item(i)).Value
      ValidaceDvojVod.Add .Range("Z" & CislaRadku.Item(i + 1)).Value
    End If
  Next i
End With

For a = 1 To ValidaceKont.Count
  For b = ValidaceVod.Count To 1 Step -1
    If a <> b And a <= ValidaceKont.Count And b <= ValidaceVod.Count Then
      If ValidaceKont.Item(a) = ValidaceKont.Item(b) And Left(ValidaceVod.Item(a), Len(CStr(ValidaceVod.Item(a))) - 2) = Left(ValidaceVod.Item(b), Len(CStr(ValidaceVod.Item(b))) - 2) And ValidaceTes.Item(a) = ValidaceTes.Item(b) Then
        ValidaceKont.Remove (b)
        ValidaceVod.Remove (b)
        ValidaceTes.Remove (b)
      End If
    End If
  Next b
Next a

For a = 1 To ValidaceDvojKont.Count - 1 Step 2
  For b = ValidaceDvojVod.Count - 1 To 1 Step -2
    If a <> b And a <= ValidaceDvojKont.Count And b <= ValidaceDvojVod.Count Then
      If ValidaceDvojKont.Item(a) = ValidaceDvojKont.Item(b) And Left(ValidaceDvojVod.Item(a), Len(CStr(ValidaceDvojVod.Item(a))) - 2) = Left(ValidaceDvojVod.Item(b), Len(CStr(ValidaceDvojVod.Item(b))) - 2) And ValidaceDvojKont.Item(a + 1) = ValidaceDvojKont.Item(b + 1) And ValidaceDvojVod.Item(a + 1) = ValidaceDvojVod.Item(b + 1) Then
        ValidaceDvojKont.Remove (b + 1)
        ValidaceDvojVod.Remove (b + 1)
        ValidaceDvojKont.Remove (b)
        ValidaceDvojVod.Remove (b)
      End If
    End If
  Next b
Next a

PocetValidaci = ValidaceKont.Count + ValidaceDvojKont.Count / 2
If PocetValidaci > 0 Then
  Set Pozadavek = Workbooks.Open("P:\TPV\NOVÁ SLOŽKA TPV\KRIMPOVACÍ NÁSTROJE\Požadavek na validaci\Formuláø a vzor Požadavek na validaci\Požadavek _validace_è.xxxx_2018_kontakt.xlsx")
  Pozadavek.Worksheets("List1").Name = "Pozadavek 1"
  Application.DisplayAlerts = False
'  Pozadavek.Worksheets("List2").Delete
'  Pozadavek.Worksheets("List3").Delete
  Application.DisplayAlerts = True
  For i = 2 To PocetValidaci
    Pozadavek.Worksheets("Pozadavek 1").Copy After:=Worksheets("Pozadavek " & i - 1)
    Pozadavek.Worksheets("Pozadavek 1 (2)").Name = "Pozadavek " & (i)
  Next i
End If

If ValidaceKont.Count > 0 Then
  For i = 1 To ValidaceKont.Count
    With Pozadavek.Worksheets("Pozadavek " & i)
      .Range("C4").Value = "è.0000/" & Year(Date)
      .Range("F3").Value = Date
      .Range("F3").HorizontalAlignment = xlHAlignCenter
      .Range("F5").Value = Application.UserName
      .Range("F5").HorizontalAlignment = xlHAlignCenter
      .Range("B8").Value = FormPozadavek.TxtCisloSvazku.Text
      .Range("B8").HorizontalAlignment = xlLeft
      .Range("C11").NumberFormat = "@"
      .Range("C11").Value = ValidaceKont.Item(i)
      .Range("C11").HorizontalAlignment = xlLeft
      .Range("C11").Font.Size = 10
      .Range("C17").Value = ValidaceVod.Item(i)
      Call TypVodice(i, 0, ValidaceVod, i)
      .Range("F19").Value = Left(Right(ValidaceVod.Item(i), 8), 2) & "," & Right(Right(ValidaceVod.Item(i), 8), 2)
      If Left(Right(ValidaceVod.Item(i), 8), 1) = "0" Then
        .Range("F19").Value = Mid(Right(ValidaceVod.Item(i), 8), 2, 1) & "," & Mid(Right(ValidaceVod.Item(i), 8), 3, 2) & " mm2"
      Else
        .Range("F19").Value = Left(Right(ValidaceVod.Item(i), 8), 2) & "," & Mid(Right(ValidaceVod.Item(i), 8), 3, 2) & " mm2"
      End If
      n = .Range("F19").Characters.Count
      .Range("F19").Characters(n, 1).Font.Superscript = True
      .CheckBoxes(1).Value = xlOff
      .Range("C24").Value = ValidaceTes.Item(i)
      .Range("C24").Font.Color = vbRed
      .Range("C36").Value = Norma
      .Activate
      ActiveWindow.ScrollRow = 1
    End With
  Next i
End If

Dim d As Integer
If ValidaceDvojKont.Count > 0 Then
    For i = 1 To ValidaceDvojKont.Count Step 2
      If i = 1 Then
        d = i
      Else
        d = i - 1
      End If
      With Pozadavek.Worksheets("Pozadavek " & ValidaceKont.Count + d)
        .Range("C4").Value = "è.0000/" & Year(Date)
        .Range("F3").Value = Date
        .Range("F3").HorizontalAlignment = xlHAlignCenter
        .Range("F5").Value = Application.UserName
        .Range("F5").HorizontalAlignment = xlHAlignCenter
        .Range("B8").Value = FormPozadavek.TxtCisloSvazku.Text
        .Range("B8").HorizontalAlignment = xlLeft
        .Range("C11").NumberFormat = "@"
        .Range("C11").Value = ValidaceDvojKont.Item(i)
        .Range("C11").HorizontalAlignment = xlLeft
        .Range("C11").Font.Size = 10
        .Range("C17").Value = ValidaceDvojVod.Item(i)
        .Range("C18").Value = ValidaceDvojVod.Item(i + 1)
        If i = 1 Then
          Call TypVodice(i, 0, ValidaceDvojVod, ValidaceKont.Count + i)
          Call TypVodice(i + 1, 1, ValidaceDvojVod, ValidaceKont.Count + i)
        Else
          Call TypVodice(i, 0, ValidaceDvojVod, ValidaceKont.Count + i - 1)
          Call TypVodice(i + 1, 1, ValidaceDvojVod, ValidaceKont.Count + i - 1)
        End If
        .Range("F19").Value = Left(Right(ValidaceDvojVod.Item(i), 8), 2) & "," & Right(Right(ValidaceDvojVod.Item(i), 8), 2)
        If Left(Right(ValidaceDvojVod.Item(i), 8), 1) = "0" Then
          .Range("F19").Value = Mid(Right(ValidaceDvojVod.Item(i), 8), 2, 1) & "," & Mid(Right(ValidaceDvojVod.Item(i), 8), 3, 2) & " mm2"
        Else
          .Range("F19").Value = Left(Right(ValidaceDvojVod.Item(i), 8), 2) & "," & Mid(Right(ValidaceDvojVod.Item(i), 8), 3, 2) & " mm2"
        End If
        n = .Range("F19").Characters.Count
        .Range("F19").Characters(n, 1).Font.Superscript = True
        .Range("F20").Value = Left(Right(ValidaceDvojVod.Item(i + 1), 8), 2) & "," & Right(Right(ValidaceDvojVod.Item(i + 1), 8), 2)
        If Left(Right(ValidaceDvojVod.Item(i + 1), 8), 1) = "0" Then
          .Range("F20").Value = Mid(Right(ValidaceDvojVod.Item(i + 1), 8), 2, 1) & "," & Mid(Right(ValidaceDvojVod.Item(i + 1), 8), 3, 2) & " mm2"
        Else
          .Range("F20").Value = Left(Right(ValidaceDvojVod.Item(i + 1), 8), 2) & "," & Mid(Right(ValidaceVod.Item(i + 1), 8), 3, 2) & " mm2"
        End If
        n = .Range("F20").Characters.Count
        .Range("F20").Characters(n, 1).Font.Superscript = True
        .CheckBoxes(1).Value = xlOn
        .Activate
        ActiveWindow.ScrollRow = 1
      End With
  Next i
End If
End Sub

Sub TypVodice(i As Long, a As Integer, Validace As Collection, list As Long)

With Pozadavek.Worksheets("Pozadavek " & list)
  If Left(Validace.Item(i), 2) = "0." Then .Range("F" & 17 + a).Value = "FL4G"
  If Left(Validace.Item(i), 2) = "01" Then .Range("F" & 17 + a).Value = "CLXPB"
  If Left(Validace.Item(i), 2) = "02" Then .Range("F" & 17 + a).Value = "CLXPC"
  If Left(Validace.Item(i), 2) = "1." Then .Range("F" & 17 + a).Value = "FLY"
  If Left(Validace.Item(i), 2) = "10" Then .Range("F" & 17 + a).Value = "FL4GY"
  If Left(Validace.Item(i), 2) = "11" Then .Range("F" & 17 + a).Value = "H05V2K"
  If Left(Validace.Item(i), 2) = "12" Then .Range("F" & 17 + a).Value = "H07V2-K"
  If Left(Validace.Item(i), 2) = "13" Then .Range("F" & 17 + a).Value = "AWM1015TEW-AWG"
  If Left(Validace.Item(i), 2) = "14" Then .Range("F" & 17 + a).Value = "FLRX-B"
  If Left(Validace.Item(i), 2) = "15" Then .Range("F" & 17 + a).Value = "H05V-K"
  If Left(Validace.Item(i), 2) = "16" Then .Range("F" & 17 + a).Value = "IEC"
  If Left(Validace.Item(i), 2) = "17" Then .Range("F" & 17 + a).Value = "MK3"
  If Left(Validace.Item(i), 2) = "18" Then .Range("F" & 17 + a).Value = "FLRYYW-B"
  If Left(Validace.Item(i), 2) = "19" Then .Range("F" & 17 + a).Value = "Vodic bez iz. pocin."
  If Left(Validace.Item(i), 3) = "2.0" And Left(Right(Validace.Item(i), 3), 1) = "0" Then .Range("F" & 17 + a).Value = "FLRY-B"
  If Left(Validace.Item(i), 3) = "2.0" And Left(Right(Validace.Item(i), 3), 1) = "1" Then .Range("F" & 17 + a).Value = "FLRY-A"
  If Left(Validace.Item(i), 4) = "2.1." Then .Range("F" & 17 + a).Value = "FLRYY-B"
  If Left(Validace.Item(i), 2) = "20" Then .Range("F" & 17 + a).Value = "FL2X"
  If Left(Validace.Item(i), 2) = "21" Then .Range("F" & 17 + a).Value = "H07VK"
  If Left(Validace.Item(i), 2) = "22" Then .Range("F" & 17 + a).Value = "H05V-K_"
  If Left(Validace.Item(i), 2) = "23" Then .Range("F" & 17 + a).Value = "H05G-K"
  If Left(Validace.Item(i), 2) = "24" Then .Range("F" & 17 + a).Value = "FLRYY"
  If Left(Validace.Item(i), 2) = "25" Then .Range("F" & 17 + a).Value = "LiY"
  If Left(Validace.Item(i), 2) = "26" Then .Range("F" & 17 + a).Value = "OLFLON FEP"
  If Left(Validace.Item(i), 2) = "27" Then .Range("F" & 17 + a).Value = "FLRNYNY"
  If Left(Validace.Item(i), 2) = "28" Then .Range("F" & 17 + a).Value = "FLR7Y"
  If Left(Validace.Item(i), 2) = "29" Then .Range("F" & 17 + a).Value = "H07RN-F"
  If Left(Validace.Item(i), 2) = "3." Then .Range("F" & 17 + a).Value = "FL2G"
  If Left(Validace.Item(i), 2) = "30" Then .Range("F" & 17 + a).Value = "FLR12Y12Y"
  If Left(Validace.Item(i), 2) = "31" Then .Range("F" & 17 + a).Value = "FLR12YNY"
  If Left(Validace.Item(i), 2) = "32" Then .Range("F" & 17 + a).Value = "FLR31YB11Y-A"
  If Left(Validace.Item(i), 4) = "33.0" Then .Range("F" & 17 + a).Value = "FLRYW-B"
  If Left(Validace.Item(i), 5) = "33.2." Or Left(Validace.Item(i), 4) = "33.3." Then .Range("F" & 17 + a).Value = "FLRYW-A"
  If Left(Validace.Item(i), 2) = "34" Then .Range("F" & 17 + a).Value = "CSZ"
  If Left(Validace.Item(i), 2) = "35" Then .Range("F" & 17 + a).Value = "WSK"
  If Left(Validace.Item(i), 2) = "36" Then .Range("F" & 17 + a).Value = "FLRYY11Y-B"
  If Left(Validace.Item(i), 2) = "37" Then .Range("F" & 17 + a).Value = "FLRW13Y"
  If Left(Validace.Item(i), 2) = "38" Then .Range("F" & 17 + a).Value = "FLWR7Y"
  If Left(Validace.Item(i), 2) = "39" Then .Range("F" & 17 + a).Value = "ACW"
  If Left(Validace.Item(i), 2) = "4." Then .Range("F" & 17 + a).Value = "FLYW"
  If Left(Validace.Item(i), 2) = "40" Then .Range("F" & 17 + a).Value = "LOSi"
  If Left(Validace.Item(i), 2) = "41" Then .Range("F" & 17 + a).Value = "FLR6Y"
  If Left(Validace.Item(i), 2) = "42" Then .Range("F" & 17 + a).Value = "T3 IR LF"
  If Left(Validace.Item(i), 2) = "43" Then .Range("F" & 17 + a).Value = "LIF6Y"
  If Left(Validace.Item(i), 2) = "44" Then .Range("F" & 17 + a).Value = "AWG"
  If Left(Validace.Item(i), 2) = "45" Then .Range("F" & 17 + a).Value = "CS-ES(flex)"
  If Left(Validace.Item(i), 2) = "46" Then .Range("F" & 17 + a).Value = "AWG UL 3398"
  If Left(Validace.Item(i), 2) = "48" Then .Range("F" & 17 + a).Value = "FLR21X"
  If Left(Validace.Item(i), 2) = "5." Then .Range("F" & 17 + a).Value = "FLYY"
  If Left(Validace.Item(i), 2) = "50" Then .Range("F" & 17 + a).Value = "T3 ZH ID"
  If Left(Validace.Item(i), 2) = "51" Then .Range("F" & 17 + a).Value = "AVS"
  If Left(Validace.Item(i), 6) = "51.003" Then .Range("F" & 17 + a).Value = "AESSX"
  If Left(Validace.Item(i), 2) = "52" Then .Range("F" & 17 + a).Value = "FLR32Y11Y"
  If Left(Validace.Item(i), 2) = "53" Then .Range("F" & 17 + a).Value = "UL(MTW)-CSA-HAR"
  If Left(Validace.Item(i), 2) = "54" Then .Range("F" & 17 + a).Value = "METTEX"
  If Left(Validace.Item(i), 2) = "55" Then .Range("F" & 17 + a).Value = "LiCy-32"
  If Left(Validace.Item(i), 2) = "56" Then .Range("F" & 17 + a).Value = "ACOME T4055"
  If Left(Validace.Item(i), 2) = "57" Then .Range("F" & 17 + a).Value = "FIAT FHT3 acc.to 91107/17"
  If Left(Validace.Item(i), 2) = "59" Then .Range("F" & 17 + a).Value = "A3Z"
  If Left(Validace.Item(i), 2) = "6." Then .Range("F" & 17 + a).Value = "FL2GY"
  If Left(Validace.Item(i), 3) = "60." And Left(Right(Validace.Item(i), 3), 1) = "0" Then .Range("F" & 17 + a).Value = "FLRY-B"
  If Left(Validace.Item(i), 3) = "60." And Left(Right(Validace.Item(i), 3), 1) = "1" Then .Range("F" & 17 + a).Value = "FLRY-A"
  If Left(Validace.Item(i), 2) = "61" Then .Range("F" & 17 + a).Value = "F3Z Renault"
  If Left(Validace.Item(i), 2) = "62" Then .Range("F" & 17 + a).Value = "FLR2X31Y"
  If Left(Validace.Item(i), 2) = "63" Then .Range("F" & 17 + a).Value = "FLR2X 125°C"
  If Left(Validace.Item(i), 2) = "64" Then .Range("F" & 17 + a).Value = "11077 SI AWG22(19sp)"
  If Left(Validace.Item(i), 2) = "65" Then .Range("F" & 17 + a).Value = "FLU7Y"
  If Left(Validace.Item(i), 2) = "66" Then .Range("F" & 17 + a).Value = "FLR13Y Arnitel C"
  If Left(Validace.Item(i), 2) = "67" Then .Range("F" & 17 + a).Value = "R1"
  If Left(Validace.Item(i), 2) = "68" Then .Range("F" & 17 + a).Value = "HIVOCAR 105-SU 0,5/2,1mm"
  If Left(Validace.Item(i), 2) = "69" Then .Range("F" & 17 + a).Value = "R2"
  If Left(Validace.Item(i), 2) = "7." Then .Range("F" & 17 + a).Value = "H03"
  If Left(Validace.Item(i), 2) = "70" Then .Range("F" & 17 + a).Value = "Leoni Mocar 150A"
  If Left(Validace.Item(i), 2) = "71" Then .Range("F" & 17 + a).Value = "Leoni Mocar 150C"
  If Left(Validace.Item(i), 2) = "72" Then .Range("F" & 17 + a).Value = "LEONI MOCAR 180G(11074)"
  If Left(Validace.Item(i), 2) = "73" Then .Range("F" & 17 + a).Value = "LEONI MOCAR 150 LAC"
  If Left(Validace.Item(i), 2) = "74" Then .Range("F" & 17 + a).Value = "LEONI MOCAR 125S Flexibilní"
  If Left(Validace.Item(i), 2) = "75" Then .Range("F" & 17 + a).Value = "LEONI MOCAR 125S"
  If Left(Validace.Item(i), 3) = "76." And Left(Right(Validace.Item(i), 3), 1) = "0" Then .Range("F" & 17 + a).Value = "FLR2X-B 150°C"
  If Left(Validace.Item(i), 3) = "76." And Left(Right(Validace.Item(i), 3), 1) = "1" Then .Range("F" & 17 + a).Value = "FLR2X-A 150°C"
  If Left(Validace.Item(i), 2) = "77" Then .Range("F" & 17 + a).Value = "A4Z"
  If Left(Validace.Item(i), 2) = "78" Then .Range("F" & 17 + a).Value = "R3"
  If Left(Validace.Item(i), 2) = "79" Then .Range("F" & 17 + a).Value = "FLW6Y"
  If Left(Validace.Item(i), 2) = "8." And Left(Right(Validace.Item(i), 3), 1) = "0" Then
    .Range("F" & 17 + a).Value = "FLR13Y-B"
  ElseIf Left(Validace.Item(i), 2) = "8." Then
    .Range("F" & 17 + a).Value = "FLR13Y-A"
  End If
  If Left(Validace.Item(i), 2) = "80" Then .Range("F" & 17 + a).Value = "LR13Y GG"
  If Left(Validace.Item(i), 2) = "81" Then .Range("F" & 17 + a).Value = "Wirinox 0,24mm"
  If Left(Validace.Item(i), 2) = "82" Then .Range("F" & 17 + a).Value = "Radox 155S FLR Anticapillary"
  If Left(Validace.Item(i), 2) = "83" Then .Range("F" & 17 + a).Value = "JUDD WIRE Anti-cap cable XLPE"
  If Left(Validace.Item(i), 2) = "84" Then .Range("F" & 17 + a).Value = "UL(AWM)-CSA-TR64"
  If Left(Validace.Item(i), 3) = "86." And Left(Right(Validace.Item(i), 3), 1) = "1" Then .Range("F" & 17 + a).Value = "FLR2X-A 150°C"
  If Left(Validace.Item(i), 3) = "86." And Left(Right(Validace.Item(i), 3), 1) = "0" Then .Range("F" & 17 + a).Value = "FLR2X-B 150°C"
  If Left(Validace.Item(i), 2) = "87" Then .Range("F" & 17 + a).Value = "LEONI MOCAR 180G(27984)"
  If Left(Validace.Item(i), 2) = "88" Then .Range("F" & 17 + a).Value = "FLR2X (flexibilní) / T150"
  If Left(Validace.Item(i), 3) = "89." And Left(Right(Validace.Item(i), 3), 1) = "1" Then .Range("F" & 17 + a).Value = "FLR9Y-A"
  If Left(Validace.Item(i), 3) = "89." And Left(Right(Validace.Item(i), 3), 1) = "0" Then .Range("F" & 17 + a).Value = "FLR9Y-B"
  If Left(Validace.Item(i), 2) = "90" Then .Range("F" & 17 + a).Value = "LEONI MOCAR 180 E"
  If Left(Validace.Item(i), 2) = "91" Then .Range("F" & 17 + a).Value = "LEONI MOCAR 260 R"
  If Left(Validace.Item(i), 2) = "92" Then .Range("F" & 17 + a).Value = "FLU2X"
  If Left(Validace.Item(i), 2) = "93" Then .Range("F" & 17 + a).Value = "FLUO"
  If Left(Validace.Item(i), 2) = "94" Then .Range("F" & 17 + a).Value = "A4F"
  If Left(Validace.Item(i), 2) = "95" Then .Range("F" & 17 + a).Value = "C4 ZH"
End With
End Sub


