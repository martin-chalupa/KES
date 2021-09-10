Attribute VB_Name = "ModulePreklad"
Option Explicit
Public Preklad As String

Public Sub GetInfo(VlozenyText As String)
  On Error GoTo Error_handler:

    Dim IE As Object
    Set IE = CreateObject("InternetExplorer.Application")
    
    Dim t As Date
    Const MAX_WAIT_SEC As Long = 5
    Preklad = vbNullString
    IE.Visible = False
    Application.Wait (Now + TimeValue("0:00:02"))
    IE.Navigate "https://translate.google.com/#view=home&op=translate&sl=auto&tl=en"

    While IE.Busy Or IE.ReadyState < 4: DoEvents: Wend

          IE.Document.querySelector("#source").Value = VlozenyText

        Dim translation As Object
        t = Timer
        Do
            On Error Resume Next
            Set translation = IE.Document.querySelector(".tlid-translation.translation")
            Preklad = translation.textContent
            On Error GoTo 0
            If Timer - t > MAX_WAIT_SEC Then Exit Do
        Loop While Preklad = vbNullString
    IE.Quit
    
    Exit Sub
Error_handler:
MsgBox "Aplikace zaznamenala problém pøi pøekladu. Je nutné dopsat pøeklad operací montáže ruènì!"
End Sub
