Attribute VB_Name = "ModDuplicity"
Option Explicit

Sub NajitDuplicity()
    
Dim Radky As Long
Dim index As Long
Dim index2 As Long


Radky = ActiveWorkbook.Sheets(3).Cells(Rows.Count, 1).End(xlUp).Row
MsgBox Radky
With ActiveWorkbook.Sheets(3)
For index = 2 To Radky
    If .Range("AA" & index) = "" Then
        For index2 = index + 1 To Radky
            If .Range("Z" & index) = .Range("Z" & index2) Then
                If .Range("AE" & index) = .Range("AE" & index2) Then
                    If KontaktX(index) = "" Then
                        If .Range("T" & index2) = "" Then
                            If .Range("S" & index) > 10 And .Range("S" & index2) > 10 Then
                                If .Range("AI" & index) = "" And .Range("AI" & index2) = "" Then
                                    If .Range("AH" & index) > 10 And .Range("AH" & index2) > 10 Then
                                        .Range("AA" & index2) = index
                                        .Range("AA" & index) = index
                                    ElseIf .Range("AH" & index) = .Range("AH" & index2) Then
                                        .Range("AA" & index2) = index
                                        .Range("AA" & index) = index
                                    End If
                                ElseIf .Range("AI" & index) = .Range("AI" & index2) Then
                                    .Range("AA" & index2) = index
                                    .Range("AA" & index) = index
                                End If
                            ElseIf .Range("S" & index) = .Range("S" & index2) Then
                                If .Range("AI" & index) = "" And .Range("AI" & index2) = "" Then
                                    If .Range("AH" & index) > 10 And .Range("AH" & index2) > 10 Then
                                        .Range("AA" & index2) = index
                                        .Range("AA" & index) = index
                                    ElseIf .Range("AH" & index) = .Range("AH" & index2) Then
                                        .Range("AA" & index2) = index
                                        .Range("AA" & index) = index
                                    End If
                                ElseIf .Range("AI" & index) = .Range("AI" & index2) Then
                                    .Range("AA" & index2) = index
                                    .Range("AA" & index) = index
                                End If
                            ElseIf .Range("S" & index) > 10 And .Range("AH" & index2) > 10 Then
                                If .Range("AI" & index) = "" And .Range("AI" & index2) = "" Then
                                    If .Range("AH" & index) > 10 And .Range("S" & index2) > 10 Then
                                        .Range("AA" & index2) = index
                                        .Range("AA" & index) = index
                                    ElseIf .Range("AH" & index) = .Range("S" & index2) Then
                                        .Range("AA" & index2) = index
                                        .Range("AA" & index) = index
                                    End If
                                End If
                            ElseIf .Range("S" & index) = .Range("AH" & index2) Then
                                If .Range("AI" & index) = "" And .Range("AI" & index2) = "" Then
                                    If .Range("AH" & index) > 10 And .Range("S" & index2) > 10 Then
                                        .Range("AA" & index2) = index
                                        .Range("AA" & index) = index
                                    ElseIf .Range("AH" & index) = .Range("S" & index2) Then
                                        .Range("AA" & index2) = index
                                        .Range("AA" & index) = index
                                    End If
                                End If
                            End If
                        ElseIf .Range("S" & index) > 10 And .Range("AH" & index2) > 10 Then
                            If .Range("AI" & index) = .Range("T" & index2) Then
                                .Range("AA" & index2) = index
                                .Range("AA" & index) = index
                            End If
                        ElseIf .Range("S" & index) = .Range("AH" & index2) Then
                            If .Range("AI" & index) = .Range("T" & index2) Then
                                .Range("AA" & index2) = index
                                .Range("AA" & index) = index
                            End If
                        End If
                    ElseIf .Range("AI" & index) = "" Then
                        If .Range("AH" & index) > 10 And .Range("AH" & index2) > 10 Then
                            If .Range("T" & index) = .Range("T" & index2) Then
                                .Range("AA" & index2) = index
                                .Range("AA" & index) = index
                            End If
                        ElseIf .Range("AH" & index) = .Range("AH" & index2) Then
                            If .Range("T" & index) = .Range("T" & index2) Then
                                .Range("AA" & index2) = index
                                .Range("AA" & index) = index
                            End If
                        ElseIf .Range("AH" & index) > 10 And .Range("S" & index2) > 10 Then
                            If .Range("T" & index) = .Range("AI" & index2) Then
                                .Range("AA" & index2) = index
                                .Range("AA" & index) = index
                            End If
                        ElseIf .Range("AH" & index) = .Range("S" & index2) Then
                            If .Range("T" & index) = .Range("AI" & index2) Then
                                .Range("AA" & index2) = index
                                .Range("AA" & index) = index
                            End If
                        End If
                    ElseIf .Range("T" & index) = .Range("T" & index2) Then
                        If .Range("AI" & index) = .Range("AI" & index2) Then
                            .Range("AA" & index2) = index
                            .Range("AA" & index) = index
                        End If
                    ElseIf .Range("T" & index) = .Range("AI" & index2) Then
                        If .Range("T" & index2) = .Range("AI" & index) Then
                            .Range("AA" & index2) = index
                            .Range("AA" & index) = index
                        End If
                    End If
                End If
            End If
        Next index2
    End If
Next index
End With

End Sub

Sub NajitDuplicityPole()
    
Dim Radky As Long
Dim index As Long
Dim index2 As Long

Dim Duplicity(42582) As Variant
Dim Vodice(42582) As Variant
Dim Delky(42582) As Variant
Dim KontaktX(42582) As Variant
Dim KontaktY(42582) As Variant
Dim OdizolX(42582) As Variant
Dim OdizolY(42582) As Variant

Radky = ActiveWorkbook.Sheets(3).Cells(Rows.Count, 1).End(xlUp).Row

With ActiveWorkbook.Sheets(3)
For index = 0 To Radky - 2
    Vodice(index) = .Range("Z" & index + 2)
    Delky(index) = .Range("AE" & index + 2)
    KontaktX(index) = .Range("T" & index + 2)
    KontaktY(index) = .Range("AI" & index + 2)
    OdizolX(index) = .Range("S" & index + 2)
    OdizolY(index) = .Range("AH" & index + 2)
Next index

For index = 0 To Radky - 3
    If IsEmpty(Duplicity(index)) = True Then
        For index2 = index + 1 To Radky - 2
            If Vodice(index) = Vodice(index2) Then
                If Delky(index) = Delky(index2) Then 'Or Delky(index) + 2 = Delky(index2) Or Delky(index) - 2 = Delky(index2) Then
                    If KontaktX(index) = "" Then
                        If KontaktX(index2) = "" Then
                            If OdizolX(index) > 10 And OdizolX(index2) > 10 Then
                                If KontaktY(index) = "" And KontaktY(index2) = "" Then
                                    If OdizolY(index) > 10 And OdizolY(index2) > 10 Then
                                        Duplicity(index2) = index
                                        Duplicity(index) = index
                                    ElseIf OdizolY(index) = OdizolY(index2) Then
                                        Duplicity(index2) = index
                                        Duplicity(index) = index
                                    End If
                                ElseIf KontaktY(index) = KontaktY(index2) Then
                                    Duplicity(index2) = index
                                    Duplicity(index) = index
                                End If
                            ElseIf OdizolX(index) = OdizolX(index2) Then
                                If KontaktY(index) = "" And KontaktY(index2) = "" Then
                                    If OdizolY(index) > 10 And OdizolY(index2) > 10 Then
                                        Duplicity(index2) = index
                                        Duplicity(index) = index
                                    ElseIf OdizolY(index) = OdizolY(index2) Then
                                        Duplicity(index2) = index
                                        Duplicity(index) = index
                                    End If
                                ElseIf KontaktY(index) = KontaktY(index2) Then
                                    Duplicity(index2) = index
                                    Duplicity(index) = index
                                End If
                            ElseIf OdizolX(index) > 10 And OdizolY(index2) > 10 Then
                                If KontaktY(index) = "" And KontaktY(index2) = "" Then
                                    If OdizolY(index) > 10 And OdizolX(index2) > 10 Then
                                        Duplicity(index2) = index
                                        Duplicity(index) = index
                                    ElseIf OdizolY(index) = OdizolX(index2) Then
                                        Duplicity(index2) = index
                                        Duplicity(index) = index
                                    End If
                                End If
                            ElseIf OdizolX(index) = OdizolY(index2) Then
                                If KontaktY(index) = "" And KontaktY(index2) = "" Then
                                    If OdizolY(index) > 10 And OdizolX(index2) > 10 Then
                                        Duplicity(index2) = index
                                        Duplicity(index) = index
                                    ElseIf OdizolY(index) = OdizolX(index2) Then
                                        Duplicity(index2) = index
                                        Duplicity(index) = index
                                    End If
                                End If
                            End If
                        ElseIf OdizolX(index) > 10 And OdizolY(index2) > 10 Then
                            If KontaktY(index) = KontaktX(index2) Then
                                Duplicity(index2) = index
                                Duplicity(index) = index
                            End If
                        ElseIf OdizolX(index) = OdizolY(index2) Then
                            If KontaktY(index) = KontaktX(index2) Then
                                Duplicity(index2) = index
                                Duplicity(index) = index
                            End If
                        End If
                    ElseIf KontaktY(index) = "" Then
                        If OdizolY(index) > 10 And OdizolY(index2) > 10 Then
                            If KontaktX(index) = KontaktX(index2) Then
                                Duplicity(index2) = index
                                Duplicity(index) = index
                            End If
                        ElseIf OdizolY(index) = OdizolY(index2) Then
                            If KontaktX(index) = KontaktX(index2) Then
                                Duplicity(index2) = index
                                Duplicity(index) = index
                            End If
                        ElseIf OdizolY(index) > 10 And OdizolX(index2) > 10 Then
                            If KontaktX(index) = KontaktY(index2) Then
                                Duplicity(index2) = index
                                Duplicity(index) = index
                            End If
                        ElseIf OdizolY(index) = OdizolX(index2) Then
                            If KontaktX(index) = KontaktY(index2) Then
                                Duplicity(index2) = index
                                Duplicity(index) = index
                            End If
                        End If
                    ElseIf KontaktX(index) = KontaktX(index2) Then
                        If KontaktY(index) = KontaktY(index2) Then
                            Duplicity(index2) = index
                            Duplicity(index) = index
                        End If
                    ElseIf KontaktX(index) = KontaktY(index2) Then
                        If KontaktX(index2) = KontaktY(index) Then
                            Duplicity(index2) = index
                            Duplicity(index) = index
                        End If
                    End If
                End If
            End If
        Next index2
    End If
Next index
For index = 0 To Radky - 2
    .Range("AA" & index + 2).Value = Duplicity(index)
Next index
End With

End Sub

Sub vymazDvojaku()

Dim index As Long
Dim Radky As Long

Radky = ActiveWorkbook.Sheets(3).Cells(Rows.Count, 1).End(xlUp).Row

With ActiveWorkbook.Sheets(3)
index = 2
Do While index <= Radky
    If .Range("AM" & index).Value = "Dvojzális" Or .Range("AM" & index).Value = "DvojzálisBK" Or .Range("AN" & index).Value = "Dvojzális" Then
        If .Range("BC" & index) = .Range("BC" & index - 1) Then
            .Rows(index).Interior.Color = vbRed
            .Rows(index - 1).Interior.Color = vbRed
            index = index + 1
        ElseIf .Range("BC" & index) = .Range("BC" & index + 1) Then
            .Rows(index).Interior.Color = vbRed
            .Rows(index + 1).Interior.Color = vbRed
            index = index + 2
        End If
    Else
        index = index + 1
    End If
Loop
End With

End Sub


Sub Doplnitcislapodskupin()

Dim Radky As Long
Dim index As Long
Dim index2 As Long
Dim k As Integer
Dim L As Integer

Radky = ActiveWorkbook.Sheets(3).Cells(Rows.Count, 1).End(xlUp).Row

With ActiveWorkbook.Sheets(3)
index = 2
k = 0
Do While index <= Radky
    index2 = 1
    k = k + 1
    Do While .Range("L" & index) = .Range("L" & index + index2)
        index2 = index2 + 1
    Loop
    .Range("M" & index) = k
    .Range("N" & index) = index2
    index = index + index2
Loop
    
End With
End Sub

Sub Test()
Dim test1 As Integer
test1 = ActiveWorkbook.Sheets(3).Range("V2").Value + 1
MsgBox test1
End Sub
