Attribute VB_Name = "VyhledaniPPAP"
Option Explicit

Public Komponent As String
Public ChybiRok As Boolean
Public ChybiPPAP As Boolean
Public Pozice As Long
Public CelyNazevSlozky As Variant
Public Prazdna As Boolean

Sub SpusteniPPAP()
Attribute SpusteniPPAP.VB_ProcData.VB_Invoke_Func = "Q\n14"

Dim Okno As Integer
Okno = MsgBox("Funkce makra:" & vbCrLf & "  - prohledává složku P:\PPAP_nakupovaných dílu pro PPAP dílù dle kusovníku exportovaného z Xpertu" & vbCrLf & "  - hledá konkrétní soubor dle klíèového slova " & Chr(34) & "KESapproved" & Chr(34) & vbCrLf & "  - v pøípadì, že je soubor nalezen, makro vytvoøí kopii tohoto souboru na ploše ve složce " _
& Chr(34) & "PPAP komponentù" & Chr(34) & vbCrLf & "  - pokud jsou ve složce dílu další podsložky oznaèené letopoèty, makro najde nejnovìjší a tu prohledává" & vbCrLf & "  - pokud ve složce dílu (popøípadì podsložce roku) existuje pouze jeden soubor, makro zkopíruje tento soubor do složky " & Chr(34) & "PPAP komponentù" & Chr(34) & vbCrLf & "  - pro vodièe makro prohledává i složky, kde je poslední trojèíslí nahrazeno " & Chr(34) & "XXX" & Chr(34) & " popøípadì " & Chr(34) & "100" & Chr(34) & vbCrLf & vbCrLf & "V souèasné dobì je pouze menšina souborù oznaèena slovem " & Chr(34) & "KESapproved" & Chr(34) & "." & vbCrLf & vbCrLf & "PRO LEPŠÍ FUNKCI MAKRA NEZAPOMEÒ PØEJMENOVAT NALEZENÉ A SCHVÁLENÉ (KES) PPAP SOUBORY - DOPLNIT " & Chr(34) & "_KESapproved" & Chr(34) & "!!!" & vbCrLf & vbCrLf & "Opravdu chceš sputit toto makro?", vbQuestion + vbYesNo, "Vyhledání PPAP")
If Okno = vbYes Then
  Call HledaniPPAP
Else
  Exit Sub
End If
End Sub


Private Sub HledaniPPAP()

Dim objFSO As Object
Dim FindKomponenty As Range
Dim FindMatchcode As Range
Dim PocetKomponent As Long
Dim Vodic As Long
Dim i As Integer
Dim SlozkaNeexistuje As Boolean
Dim Uzivatel As String
Dim ChybiDil As Boolean
Dim Vytvoreni As String

Set objFSO = CreateObject("Scripting.FileSystemObject")

Application.DisplayAlerts = False
On Error Resume Next
ActiveWorkbook.Sheets("List2").Delete
ActiveWorkbook.Sheets("List3").Delete
On Error GoTo 0
Application.DisplayAlerts = True

If ActiveWorkbook.Sheets.Count = 1 Then
  ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = "PPAP komponentù"
End If
ActiveWorkbook.Sheets("PPAP komponentù").Cells.Clear

Set FindKomponenty = ActiveWorkbook.Sheets(1).Range("A1:AZ1").Find(What:="Komponenty", LookIn:=xlValues)
If FindKomponenty Is Nothing Then
  Set FindKomponenty = ActiveWorkbook.Sheets(1).Range("A1:AZ1").Find(What:="Èíslo dílu", LookIn:=xlValues)
End If

Set FindMatchcode = ActiveWorkbook.Sheets(1).Range("A1:AZ1").Find(What:="Matchcode", LookIn:=xlValues)
If FindMatchcode Is Nothing Then
  Set FindMatchcode = ActiveWorkbook.Sheets(1).Range("A1:AZ1").Find(What:="Oznaèení", LookIn:=xlValues)
End If

FindKomponenty.EntireColumn.Copy (ActiveWorkbook.Sheets(2).Range("A1"))
FindMatchcode.EntireColumn.Copy (ActiveWorkbook.Sheets(2).Range("B1"))
ActiveWorkbook.Sheets(2).Range("A:B").RemoveDuplicates Columns:=1
ActiveWorkbook.Sheets(2).Range("A1").EntireColumn.AutoFit
ActiveWorkbook.Sheets(2).Range("B1").EntireColumn.AutoFit

ActiveWorkbook.Sheets(2).Range("C1").Value = "PPAP dílu nalezen/nenalezen"
ActiveWorkbook.Sheets(2).Range("C1").Font.Bold = True
ActiveWorkbook.Sheets(2).Range("D1").Value = "Odkaz na PPAP dílu"
ActiveWorkbook.Sheets(2).Range("D1").Font.Bold = True
ActiveWorkbook.Sheets(2).Range("E1").Value = "Nutnost zásahu uživatele"
ActiveWorkbook.Sheets(2).Range("E1").Font.Bold = True
ActiveWorkbook.Sheets(2).Range("F1").Value = "Odkaz na pùvodní složku dílu"
ActiveWorkbook.Sheets(2).Range("F1").Font.Bold = True

PocetKomponent = ActiveWorkbook.Sheets(2).Cells(Rows.Count, 1).End(xlUp).Row

Uzivatel = StripAccent(Right(Application.UserName, Len(Application.UserName) - InStr(1, Application.UserName, " ", vbBinaryCompare)))

Vytvoreni = Format(Now, "yyyy-mm-dd hh-mm-ss")

On Error Resume Next
If Not objFSO.FolderExists("C:\Users\" & Uzivatel & "\Desktop\PPAP komponentù") Then
  objFSO.CreateFolder "C:\Users\" & Uzivatel & "\Desktop\PPAP komponentù"
End If
If Err.Number > 0 Then SlozkaNeexistuje = True
On Error GoTo 0

On Error Resume Next
If Not objFSO.FolderExists("C:\Users\" & Uzivatel & "\Desktop\PPAP komponentù\" & Vytvoreni) Then
  objFSO.CreateFolder "C:\Users\" & Uzivatel & "\Desktop\PPAP komponentù\" & Vytvoreni
End If
If Err.Number > 0 Then SlozkaNeexistuje = True
On Error GoTo 0

For i = 2 To PocetKomponent
  Komponent = ActiveWorkbook.Sheets(2).Cells(i, 1).Value
  Vodic = InStr(1, ActiveWorkbook.Sheets(2).Cells(i, 2).Value, "VOD", vbTextCompare)
  ChybiPPAP = True
  Prazdna = False
  ChybiDil = False
    If objFSO.FolderExists("P:\PPAP_nakupovane dily\" & Komponent) Then
      Call KontrolaPPAP(Komponent, SlozkaNeexistuje, i, Uzivatel, Vytvoreni)
    ElseIf Vodic > 0 Then
      Komponent = Left(ActiveWorkbook.Sheets(2).Cells(i, 1).Value, Len(ActiveWorkbook.Sheets(2).Cells(i, 1).Value) - 3)
      Komponent = Komponent & "xxx"
      If objFSO.FolderExists("P:\PPAP_nakupovane dily\" & Komponent) Then
        Call KontrolaPPAP(Komponent, SlozkaNeexistuje, i, Uzivatel, Vytvoreni)
      Else
        Komponent = Left(ActiveWorkbook.Sheets(2).Cells(i, 1).Value, Len(ActiveWorkbook.Sheets(2).Cells(i, 1).Value) - 3)
        Komponent = Komponent & "100"
        If objFSO.FolderExists("P:\PPAP_nakupovane dily\" & Komponent) Then
          Call KontrolaPPAP(Komponent, SlozkaNeexistuje, i, Uzivatel, Vytvoreni)
        Else
          ChybiDil = True
        End If
      End If
    Else
      ChybiDil = True
    End If
    
    If ChybiPPAP = True And Prazdna = False And ChybiDil = False And ChybiRok = False Then
      ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Ve složce (viz hyperlink) nebyl nalezen žádný soubor s klíèovým slovem " & Chr(34) & "KESapproved" & Chr(34) & "."
      ActiveWorkbook.Sheets(2).Cells(i, 5).Value = "Pokud složka obsahuje PPAP dílu, pøejmenuj tento soubor, aby obsahoval slovo " & Chr(34) & "_KESapproved" & Chr(34) & "."
      ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice), TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice)
      ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbRed
    ElseIf ChybiPPAP = True And Prazdna = True And ChybiDil = False Then
      If IsEmpty(CelyNazevSlozky) Then
        ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Aktuální složka (" & Komponent & ") neobsahuje žádné soubory."
        ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent, TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent
        ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbRed
      Else
        ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Aktuální složka (" & CelyNazevSlozky(Pozice) & ") neobsahuje žádné soubory. Pokud je tato složka prázdná, tak ji smaž."
        ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice), TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice)
        ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbRed
      End If
    ElseIf ChybiDil = True Then
      ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Složka dílu nebyla nalezena."
      ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbRed
    ElseIf ChybiPPAP = True And ChybiRok = True Then
      If IsEmpty(CelyNazevSlozky) Then
        ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "PPAP dílu nebyl nalezen. Ve složce dílu byla nalezena složka s jiným názvem než je letopoèet."
        ActiveWorkbook.Sheets(2).Cells(i, 5).Value = "Pokud složka obsahuje PPAP dílu, pøejmenuj tento soubor, aby obsahoval slovo " & Chr(34) & "_KESapproved" & Chr(34) & " a pøesuò soubor do složky " & Komponent & "."
        ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent, TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent
        ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbRed
      Else
        ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "PPAP dílu nebyl nalezen. Ve složce nebyla nalezena žádná složka."
        ActiveWorkbook.Sheets(2).Cells(i, 5).Value = "Pokud složka obsahuje PPAP dílu, pøejmenuj tento soubor, aby obsahoval slovo " & Chr(34) & "_KESapproved" & Chr(34) & " a pøesuò soubor do složky " & Komponent & "."
        ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent, TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent
        ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbRed
      End If
    End If
Next i

If SlozkaNeexistuje = True Then MsgBox ("Složka " & Chr(34) & "PPAP komponentù" & Chr(34) & " nemohla být vytvoøena! Nalezené PPAP dílù nebyly nikam zkopírovány!")

ActiveWorkbook.Sheets(2).Range("C1").EntireColumn.AutoFit
ActiveWorkbook.Sheets(2).Range("D1").EntireColumn.AutoFit
ActiveWorkbook.Sheets(2).Range("E1").EntireColumn.AutoFit
ActiveWorkbook.Sheets(2).Range("F1").EntireColumn.AutoFit

Worksheets(2).Activate
Exit Sub
End Sub
Private Sub KontrolaPPAP(Komponent As String, SlozkaNeexistuje As Boolean, i As Integer, Uzivatel As String, Vytvoreni As String)

  Dim objFSO As Object
  Dim objSoubory As Object
  Dim objSoubor As Object
  Dim objSlozka As Object
  Dim objSlozky As Object
  Dim a As Integer
  Dim b As Integer
  Dim strSlozka As String
  Dim Rok As Long
  Dim strRok As String
  Dim AktualniSlozka As Variant
  Dim RokSlozky As Variant
  Dim PocetSlozek As Long
  Dim PocetSouboru As Long
  Dim objAktualni As Object
  Dim objAktSoubor As Object
  
  Set objFSO = CreateObject("Scripting.FileSystemObject")
   
  Set objSoubory = objFSO.GetFolder("P:\PPAP_nakupovane dily\" & Komponent).Files
  PocetSouboru = objSoubory.Count
  
  If PocetSouboru > 0 Then
  
    For Each objSoubor In objSoubory
      If ChybiPPAP = True Then
        If objSoubor.Name Like "*.pdf" Then
          If InStr(1, objSoubor.Name, "KESapproved", vbTextCompare) > 0 Or InStr(1, objSoubor.Name, "KES approved", vbTextCompare) > 0 Then
            If SlozkaNeexistuje = False Then
              objFSO.CopyFile "P:\PPAP_nakupovane dily\" & Komponent & "\" & objSoubor.Name, "C:\Users\" & Uzivatel & "\Desktop\PPAP komponentù\" & Vytvoreni & "\PPAP_" & Komponent & ".pdf"
              ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Soubor byl zkopírován na plochu do složky " & Chr(34) & "PPAP komponentù\" & Vytvoreni & Chr(34) & "."
              ActiveWorkbook.Sheets(2).Cells(i, 4).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 4), Address:="C:\Users\" & Uzivatel & "\Desktop\PPAP komponentù\" & Vytvoreni & "\PPAP_" & Komponent & ".pdf", TextToDisplay:="PPAP_" & Komponent & ".pdf"
              ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbGreen
              ChybiPPAP = False
            Else
              ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Soubor nebyl zkopírován!"
              ActiveWorkbook.Sheets(2).Cells(i, 4).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 4), Address:="P:\PPAP_nakupovane dily\" & Komponent & "\" & objSoubor.Name, TextToDisplay:=objSoubor.Name
              ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent, TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent
              ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbGreen
              ChybiPPAP = False
            End If
          End If
        End If
      End If
    Next
    
  End If
  
  If PocetSouboru = 1 And ChybiPPAP = True Then
  
    For Each objSoubor In objSoubory
      If ChybiPPAP = True Then
        If objSoubor.Name Like "*.pdf" Then
          If SlozkaNeexistuje = False Then
            objFSO.CopyFile "P:\PPAP_nakupovane dily\" & Komponent & "\" & objSoubor.Name, "C:\Users\" & Uzivatel & "\Desktop\PPAP komponentù\" & Vytvoreni & "\Zkontroluj PPAP_" & Komponent & ".pdf"
            ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Byl nalezen pouze jeden soubor, který byl zkopírován na plochu do složky " & Chr(34) & "PPAP komponentù\" & Vytvoreni & Chr(34) & "."
            ActiveWorkbook.Sheets(2).Cells(i, 4).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 4), Address:="C:\Users\" & Uzivatel & "\Desktop\PPAP komponentù\" & Vytvoreni & "\" & "Zkontroluj PPAP_" & Komponent & ".pdf", TextToDisplay:="Zkontroluj PPAP_" & Komponent & ".pdf"
            ActiveWorkbook.Sheets(2).Cells(i, 5).Value = "Zkontroluj, zda tento soubor je skuteènì PPAP daného dílu. Pokud ano, pøejmenuj tento soubor, aby obsahoval slovo " & Chr(34) & "_KESapproved" & Chr(34) & "."
            ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent, TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent
            ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbYellow
            ChybiPPAP = False
          Else
            ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Byl nalezen pouze jeden soubor, který nebyl nikam zkopírován!"
            ActiveWorkbook.Sheets(2).Cells(i, 4).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 4), Address:="P:\PPAP_nakupovane dily\" & Komponent & "\" & objSoubor.Name, TextToDisplay:=objSoubor.Name
            ActiveWorkbook.Sheets(2).Cells(i, 5).Value = "Zkontroluj, zda tento soubor je skuteènì PPAP daného dílu. Pokud ano, pøejmenuj tento soubor, aby obsahoval slovo " & Chr(34) & "_KESapproved" & Chr(34) & "."
            ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent, TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent
            ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbYellow
            ChybiPPAP = False
          End If
        End If
      End If
    Next
    
  End If
    
  If PocetSouboru = 0 Or ChybiPPAP = True Then
    Set objSlozky = objFSO.GetFolder("P:\PPAP_nakupovane dily\" & Komponent).Subfolders
    PocetSlozek = objSlozky.Count
    If PocetSlozek > 0 Then
      ReDim CelyNazevSlozky(1 To PocetSlozek)
      ReDim RokSlozky(1 To PocetSlozek)
      a = 1
      b = 1
      ChybiRok = True
      
      For Each objSlozka In objSlozky
        strSlozka = objSlozka.Name
        CelyNazevSlozky(b) = strSlozka
        b = b + 1
        If InStr(strSlozka, "20") > 0 Then
          strRok = Mid(strSlozka, InStr(strSlozka, "20"), 4)
          ChybiRok = False
          Rok = CInt(strRok)
          RokSlozky(a) = Rok
          a = a + 1
        ElseIf InStr(strSlozka, "19") > 0 Then
          strRok = Mid(strSlozka, InStr(strSlozka, "19"), 4)
          ChybiRok = False
          Rok = CInt(strRok)
          RokSlozky(a) = Rok
          a = a + 1
        Else
          RokSlozky(a) = "0"
          a = a + 1
        End If
      Next
      
      If ChybiRok = False Then
        AktualniSlozka = RokSlozky
      
        Pozice = Application.WorksheetFunction.Match(Application.Max(AktualniSlozka), AktualniSlozka, 0)
        Set objAktualni = objFSO.GetFolder("P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice)).Files

        If objAktualni.Count = 0 Then
          Prazdna = True
        End If
        
        For Each objAktSoubor In objAktualni
          If ChybiPPAP = True Then
            If objAktSoubor.Name Like "*.pdf" Then
              If InStr(1, objAktSoubor.Name, "KESapproved", vbTextCompare) > 0 Or InStr(1, objAktSoubor.Name, "KES approved", vbTextCompare) > 0 Then
                If SlozkaNeexistuje = False Then
                  objFSO.CopyFile "P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice) & "\" & objAktSoubor.Name, "C:\Users\" & Uzivatel & "\Desktop\PPAP komponentù\" & Vytvoreni & "\PPAP_" & Komponent & ".pdf"
                  ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Soubor byl zkopírován na plochu do složky " & Chr(34) & "PPAP komponentù\" & Vytvoreni & Chr(34) & "."
                  ActiveWorkbook.Sheets(2).Cells(i, 4).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 4), Address:="C:\Users\" & Uzivatel & "\Desktop\PPAP komponentù\" & Vytvoreni & "\PPAP_" & Komponent & ".pdf", TextToDisplay:="PPAP_" & Komponent & ".pdf"
                  ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbGreen
                  ChybiPPAP = False
                Else
                  ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Soubor nebyl zkopírován!"
                  ActiveWorkbook.Sheets(2).Cells(i, 4).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 4), Address:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice) & "\" & objAktSoubor.Name, TextToDisplay:=objAktSoubor.Name
                  ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice), TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice)
                  ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbGreen
                  ChybiPPAP = False
                End If
              End If
            End If
          End If
        Next
        
        If objAktualni.Count = 1 And ChybiPPAP = True Then
        
          For Each objAktSoubor In objAktualni
            If ChybiPPAP = True Then
              If objAktSoubor.Name Like "*.pdf" Then
                If SlozkaNeexistuje = False Then
                  objFSO.CopyFile "P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice) & "\" & objAktSoubor.Name, "C:\Users\" & Uzivatel & "\Desktop\PPAP komponentù\" & Vytvoreni & "\Zkontroluj PPAP_" & Komponent & ".pdf"
                  ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Byl nalezen pouze jeden soubor, který byl zkopírován na plochu do složky " & Chr(34) & "PPAP komponentù\" & Vytvoreni & Chr(34) & "."
                  ActiveWorkbook.Sheets(2).Cells(i, 4).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 4), Address:="C:\Users\" & Uzivatel & "\Desktop\PPAP komponentù\" & Vytvoreni & "\Zkontroluj PPAP_" & Komponent & ".pdf", TextToDisplay:="Zkontroluj PPAP_" & Komponent & ".pdf"
                  ActiveWorkbook.Sheets(2).Cells(i, 5).Value = "Zkontroluj, zda tento soubor je skuteènì PPAP daného dílu. Pokud ano, pøejmenuj tento soubor, aby obsahoval slovo " & Chr(34) & "_KESapproved" & Chr(34) & "."
                  ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice), TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice)
                  ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbYellow
                  ChybiPPAP = False
                Else
                  ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Byl nalezen pouze jeden soubor, který nebyl nikam zkopírován!"
                  ActiveWorkbook.Sheets(2).Cells(i, 4).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 4), Address:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice) & "\" & objAktSoubor.Name, TextToDisplay:=objAktSoubor.Name
                  ActiveWorkbook.Sheets(2).Cells(i, 5).Value = "Zkontroluj, zda tento soubor je skuteènì PPAP daného dílu. Pokud ano, pøejmenuj tento soubor, aby obsahoval slovo " & Chr(34) & "_KESapproved" & Chr(34) & "."
                  ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice), TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice)
                  ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbYellow
                  ChybiPPAP = False
                End If
              End If
            End If
          Next
        
        End If
                
      End If
    Else
      If Not IsEmpty(CelyNazevSlozky) Then
        CelyNazevSlozky = Empty
      End If
      Prazdna = True
    End If
  End If
End Sub
Function StripAccent(strstring As String)

  Dim Diakritika As String
  Dim BezDiakritiky As String
  Dim i As Integer
  Const AccChars = "ìšèøžýáíéóïòúùÌŠÈØŽÝÁÍÉÓÏÒÚÙQWERTZUIOPASDFGHJKLYXCVBNM"
  Const RegChars = "escrzyaieodtnuuescrzyaieodtnuuqwertzuiopasdfghjklyxcvbnm"
  For i = 1 To Len(AccChars)
    Diakritika = Mid(AccChars, i, 1)
    BezDiakritiky = Mid(RegChars, i, 1)
    strstring = Replace(strstring, Diakritika, BezDiakritiky)
  Next
  StripAccent = strstring
End Function

