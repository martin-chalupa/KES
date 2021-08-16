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
Okno = MsgBox("Funkce makra:" & vbCrLf & "  - prohled�v� slo�ku P:\PPAP_nakupovan�ch d�lu pro PPAP d�l� dle kusovn�ku exportovan�ho z Xpertu" & vbCrLf & "  - hled� konkr�tn� soubor dle kl��ov�ho slova " & Chr(34) & "KESapproved" & Chr(34) & vbCrLf & "  - v p��pad�, �e je soubor nalezen, makro vytvo�� kopii tohoto souboru na plo�e ve slo�ce " _
& Chr(34) & "PPAP komponent�" & Chr(34) & vbCrLf & "  - pokud jsou ve slo�ce d�lu dal�� podslo�ky ozna�en� letopo�ty, makro najde nejnov�j�� a tu prohled�v�" & vbCrLf & "  - pokud ve slo�ce d�lu (pop��pad� podslo�ce roku) existuje pouze jeden soubor, makro zkop�ruje tento soubor do slo�ky " & Chr(34) & "PPAP komponent�" & Chr(34) & vbCrLf & "  - pro vodi�e makro prohled�v� i slo�ky, kde je posledn� troj��sl� nahrazeno " & Chr(34) & "XXX" & Chr(34) & " pop��pad� " & Chr(34) & "100" & Chr(34) & vbCrLf & vbCrLf & "V sou�asn� dob� je pouze men�ina soubor� ozna�ena slovem " & Chr(34) & "KESapproved" & Chr(34) & "." & vbCrLf & vbCrLf & "PRO LEP�� FUNKCI MAKRA NEZAPOME� P�EJMENOVAT NALEZEN� A SCHV�LEN� (KES) PPAP SOUBORY - DOPLNIT " & Chr(34) & "_KESapproved" & Chr(34) & "!!!" & vbCrLf & vbCrLf & "Opravdu chce� sputit toto makro?", vbQuestion + vbYesNo, "Vyhled�n� PPAP")
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
  ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = "PPAP komponent�"
End If
ActiveWorkbook.Sheets("PPAP komponent�").Cells.Clear

Set FindKomponenty = ActiveWorkbook.Sheets(1).Range("A1:AZ1").Find(What:="Komponenty", LookIn:=xlValues)
If FindKomponenty Is Nothing Then
  Set FindKomponenty = ActiveWorkbook.Sheets(1).Range("A1:AZ1").Find(What:="��slo d�lu", LookIn:=xlValues)
End If

Set FindMatchcode = ActiveWorkbook.Sheets(1).Range("A1:AZ1").Find(What:="Matchcode", LookIn:=xlValues)
If FindMatchcode Is Nothing Then
  Set FindMatchcode = ActiveWorkbook.Sheets(1).Range("A1:AZ1").Find(What:="Ozna�en�", LookIn:=xlValues)
End If

FindKomponenty.EntireColumn.Copy (ActiveWorkbook.Sheets(2).Range("A1"))
FindMatchcode.EntireColumn.Copy (ActiveWorkbook.Sheets(2).Range("B1"))
ActiveWorkbook.Sheets(2).Range("A:B").RemoveDuplicates Columns:=1
ActiveWorkbook.Sheets(2).Range("A1").EntireColumn.AutoFit
ActiveWorkbook.Sheets(2).Range("B1").EntireColumn.AutoFit

ActiveWorkbook.Sheets(2).Range("C1").Value = "PPAP d�lu nalezen/nenalezen"
ActiveWorkbook.Sheets(2).Range("C1").Font.Bold = True
ActiveWorkbook.Sheets(2).Range("D1").Value = "Odkaz na PPAP d�lu"
ActiveWorkbook.Sheets(2).Range("D1").Font.Bold = True
ActiveWorkbook.Sheets(2).Range("E1").Value = "Nutnost z�sahu u�ivatele"
ActiveWorkbook.Sheets(2).Range("E1").Font.Bold = True
ActiveWorkbook.Sheets(2).Range("F1").Value = "Odkaz na p�vodn� slo�ku d�lu"
ActiveWorkbook.Sheets(2).Range("F1").Font.Bold = True

PocetKomponent = ActiveWorkbook.Sheets(2).Cells(Rows.Count, 1).End(xlUp).Row

Uzivatel = StripAccent(Right(Application.UserName, Len(Application.UserName) - InStr(1, Application.UserName, " ", vbBinaryCompare)))

Vytvoreni = Format(Now, "yyyy-mm-dd hh-mm-ss")

On Error Resume Next
If Not objFSO.FolderExists("C:\Users\" & Uzivatel & "\Desktop\PPAP komponent�") Then
  objFSO.CreateFolder "C:\Users\" & Uzivatel & "\Desktop\PPAP komponent�"
End If
If Err.Number > 0 Then SlozkaNeexistuje = True
On Error GoTo 0

On Error Resume Next
If Not objFSO.FolderExists("C:\Users\" & Uzivatel & "\Desktop\PPAP komponent�\" & Vytvoreni) Then
  objFSO.CreateFolder "C:\Users\" & Uzivatel & "\Desktop\PPAP komponent�\" & Vytvoreni
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
      ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Ve slo�ce (viz hyperlink) nebyl nalezen ��dn� soubor s kl��ov�m slovem " & Chr(34) & "KESapproved" & Chr(34) & "."
      ActiveWorkbook.Sheets(2).Cells(i, 5).Value = "Pokud slo�ka obsahuje PPAP d�lu, p�ejmenuj tento soubor, aby obsahoval slovo " & Chr(34) & "_KESapproved" & Chr(34) & "."
      ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice), TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice)
      ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbRed
    ElseIf ChybiPPAP = True And Prazdna = True And ChybiDil = False Then
      If IsEmpty(CelyNazevSlozky) Then
        ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Aktu�ln� slo�ka (" & Komponent & ") neobsahuje ��dn� soubory."
        ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent, TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent
        ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbRed
      Else
        ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Aktu�ln� slo�ka (" & CelyNazevSlozky(Pozice) & ") neobsahuje ��dn� soubory. Pokud je tato slo�ka pr�zdn�, tak ji sma�."
        ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice), TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice)
        ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbRed
      End If
    ElseIf ChybiDil = True Then
      ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Slo�ka d�lu nebyla nalezena."
      ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbRed
    ElseIf ChybiPPAP = True And ChybiRok = True Then
      If IsEmpty(CelyNazevSlozky) Then
        ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "PPAP d�lu nebyl nalezen. Ve slo�ce d�lu byla nalezena slo�ka s jin�m n�zvem ne� je letopo�et."
        ActiveWorkbook.Sheets(2).Cells(i, 5).Value = "Pokud slo�ka obsahuje PPAP d�lu, p�ejmenuj tento soubor, aby obsahoval slovo " & Chr(34) & "_KESapproved" & Chr(34) & " a p�esu� soubor do slo�ky " & Komponent & "."
        ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent, TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent
        ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbRed
      Else
        ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "PPAP d�lu nebyl nalezen. Ve slo�ce nebyla nalezena ��dn� slo�ka."
        ActiveWorkbook.Sheets(2).Cells(i, 5).Value = "Pokud slo�ka obsahuje PPAP d�lu, p�ejmenuj tento soubor, aby obsahoval slovo " & Chr(34) & "_KESapproved" & Chr(34) & " a p�esu� soubor do slo�ky " & Komponent & "."
        ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent, TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent
        ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbRed
      End If
    End If
Next i

If SlozkaNeexistuje = True Then MsgBox ("Slo�ka " & Chr(34) & "PPAP komponent�" & Chr(34) & " nemohla b�t vytvo�ena! Nalezen� PPAP d�l� nebyly nikam zkop�rov�ny!")

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
              objFSO.CopyFile "P:\PPAP_nakupovane dily\" & Komponent & "\" & objSoubor.Name, "C:\Users\" & Uzivatel & "\Desktop\PPAP komponent�\" & Vytvoreni & "\PPAP_" & Komponent & ".pdf"
              ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Soubor byl zkop�rov�n na plochu do slo�ky " & Chr(34) & "PPAP komponent�\" & Vytvoreni & Chr(34) & "."
              ActiveWorkbook.Sheets(2).Cells(i, 4).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 4), Address:="C:\Users\" & Uzivatel & "\Desktop\PPAP komponent�\" & Vytvoreni & "\PPAP_" & Komponent & ".pdf", TextToDisplay:="PPAP_" & Komponent & ".pdf"
              ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbGreen
              ChybiPPAP = False
            Else
              ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Soubor nebyl zkop�rov�n!"
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
            objFSO.CopyFile "P:\PPAP_nakupovane dily\" & Komponent & "\" & objSoubor.Name, "C:\Users\" & Uzivatel & "\Desktop\PPAP komponent�\" & Vytvoreni & "\Zkontroluj PPAP_" & Komponent & ".pdf"
            ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Byl nalezen pouze jeden soubor, kter� byl zkop�rov�n na plochu do slo�ky " & Chr(34) & "PPAP komponent�\" & Vytvoreni & Chr(34) & "."
            ActiveWorkbook.Sheets(2).Cells(i, 4).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 4), Address:="C:\Users\" & Uzivatel & "\Desktop\PPAP komponent�\" & Vytvoreni & "\" & "Zkontroluj PPAP_" & Komponent & ".pdf", TextToDisplay:="Zkontroluj PPAP_" & Komponent & ".pdf"
            ActiveWorkbook.Sheets(2).Cells(i, 5).Value = "Zkontroluj, zda tento soubor je skute�n� PPAP dan�ho d�lu. Pokud ano, p�ejmenuj tento soubor, aby obsahoval slovo " & Chr(34) & "_KESapproved" & Chr(34) & "."
            ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent, TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent
            ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbYellow
            ChybiPPAP = False
          Else
            ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Byl nalezen pouze jeden soubor, kter� nebyl nikam zkop�rov�n!"
            ActiveWorkbook.Sheets(2).Cells(i, 4).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 4), Address:="P:\PPAP_nakupovane dily\" & Komponent & "\" & objSoubor.Name, TextToDisplay:=objSoubor.Name
            ActiveWorkbook.Sheets(2).Cells(i, 5).Value = "Zkontroluj, zda tento soubor je skute�n� PPAP dan�ho d�lu. Pokud ano, p�ejmenuj tento soubor, aby obsahoval slovo " & Chr(34) & "_KESapproved" & Chr(34) & "."
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
                  objFSO.CopyFile "P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice) & "\" & objAktSoubor.Name, "C:\Users\" & Uzivatel & "\Desktop\PPAP komponent�\" & Vytvoreni & "\PPAP_" & Komponent & ".pdf"
                  ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Soubor byl zkop�rov�n na plochu do slo�ky " & Chr(34) & "PPAP komponent�\" & Vytvoreni & Chr(34) & "."
                  ActiveWorkbook.Sheets(2).Cells(i, 4).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 4), Address:="C:\Users\" & Uzivatel & "\Desktop\PPAP komponent�\" & Vytvoreni & "\PPAP_" & Komponent & ".pdf", TextToDisplay:="PPAP_" & Komponent & ".pdf"
                  ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbGreen
                  ChybiPPAP = False
                Else
                  ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Soubor nebyl zkop�rov�n!"
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
                  objFSO.CopyFile "P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice) & "\" & objAktSoubor.Name, "C:\Users\" & Uzivatel & "\Desktop\PPAP komponent�\" & Vytvoreni & "\Zkontroluj PPAP_" & Komponent & ".pdf"
                  ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Byl nalezen pouze jeden soubor, kter� byl zkop�rov�n na plochu do slo�ky " & Chr(34) & "PPAP komponent�\" & Vytvoreni & Chr(34) & "."
                  ActiveWorkbook.Sheets(2).Cells(i, 4).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 4), Address:="C:\Users\" & Uzivatel & "\Desktop\PPAP komponent�\" & Vytvoreni & "\Zkontroluj PPAP_" & Komponent & ".pdf", TextToDisplay:="Zkontroluj PPAP_" & Komponent & ".pdf"
                  ActiveWorkbook.Sheets(2).Cells(i, 5).Value = "Zkontroluj, zda tento soubor je skute�n� PPAP dan�ho d�lu. Pokud ano, p�ejmenuj tento soubor, aby obsahoval slovo " & Chr(34) & "_KESapproved" & Chr(34) & "."
                  ActiveWorkbook.Sheets(2).Cells(i, 6).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 6), Address:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice), TextToDisplay:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice)
                  ActiveWorkbook.Sheets(2).Cells(i, 4).Interior.Color = vbYellow
                  ChybiPPAP = False
                Else
                  ActiveWorkbook.Sheets(2).Cells(i, 3).Value = "Byl nalezen pouze jeden soubor, kter� nebyl nikam zkop�rov�n!"
                  ActiveWorkbook.Sheets(2).Cells(i, 4).Hyperlinks.Add Anchor:=ActiveWorkbook.Sheets(2).Cells(i, 4), Address:="P:\PPAP_nakupovane dily\" & Komponent & "\" & CelyNazevSlozky(Pozice) & "\" & objAktSoubor.Name, TextToDisplay:=objAktSoubor.Name
                  ActiveWorkbook.Sheets(2).Cells(i, 5).Value = "Zkontroluj, zda tento soubor je skute�n� PPAP dan�ho d�lu. Pokud ano, p�ejmenuj tento soubor, aby obsahoval slovo " & Chr(34) & "_KESapproved" & Chr(34) & "."
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
  Const AccChars = "�������������̊�؎�����ύ���QWERTZUIOPASDFGHJKLYXCVBNM"
  Const RegChars = "escrzyaieodtnuuescrzyaieodtnuuqwertzuiopasdfghjklyxcvbnm"
  For i = 1 To Len(AccChars)
    Diakritika = Mid(AccChars, i, 1)
    BezDiakritiky = Mid(RegChars, i, 1)
    strstring = Replace(strstring, Diakritika, BezDiakritiky)
  Next
  StripAccent = strstring
End Function

Function OdstraneniDvojtecek(Vytvoreni As String)

  Dim Dvojtecky As String
  Dim BezDvojtecek As String
  Dim i As Integer
  Const AccChars = ":"
  Const RegChars = "-"
  For i = 1 To Len(AccChars)
    Dvojtecky = Mid(AccChars, i, 1)
    BezDvojtecek = Mid(RegChars, i, 1)
    Vytvoreni = Replace(Vytvoreni, Dvojtecky, BezDvojtecek)
  Next
  OdstraneniDvojtecek = Vytvoreni
End Function

