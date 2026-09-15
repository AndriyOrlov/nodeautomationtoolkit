Attribute VB_Name = "ClosedNoticePipeline"
Option Explicit

Private Function StemWord(ByVal w As String) As String
    Dim lw As String
    lw = LCase(Trim(w))
    If Len(lw) < 3 Then
        StemWord = lw
        Exit Function
    End If

    If Left(lw, 4) = ChrW(1073) & ChrW(1088) & ChrW(1080) & ChrW(1075) Then
        StemWord = ChrW(1073) & ChrW(1088) & ChrW(1080) & ChrW(1075)
        Exit Function
    End If
    If Left(lw, 5) = ChrW(1084) & ChrW(1077) & ChrW(1093) & ChrW(1072) & ChrW(1085) Then
        StemWord = ChrW(1084) & ChrW(1077) & ChrW(1093) & ChrW(1072) & ChrW(1085)
        Exit Function
    End If
    If Left(lw, 6) = ChrW(1076) & ChrW(1077) & ChrW(1089) & ChrW(1072) & ChrW(1085) & ChrW(1090) Then
        StemWord = ChrW(1076) & ChrW(1077) & ChrW(1089) & ChrW(1072) & ChrW(1085) & ChrW(1090)
        Exit Function
    End If
    If Left(lw, 5) = ChrW(1072) & ChrW(1088) & ChrW(1090) & ChrW(1080) & ChrW(1083) Then
        StemWord = ChrW(1072) & ChrW(1088) & ChrW(1090) & ChrW(1080) & ChrW(1083)
        Exit Function
    End If
    If Left(lw, 6) = ChrW(1073) & ChrW(1072) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1100) Then
        StemWord = ChrW(1073) & ChrW(1072) & ChrW(1090)
        Exit Function
    End If
    If Left(lw, 8) = ChrW(1073) & ChrW(1072) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1100) & ChrW(1086) & ChrW(1085) Then
        StemWord = ChrW(1073) & ChrW(1072) & ChrW(1090)
        Exit Function
    End If
    If Left(lw, 5) = ChrW(1094) & ChrW(1077) & ChrW(1085) & ChrW(1090) & ChrW(1088) Then
        StemWord = ChrW(1094) & ChrW(1077) & ChrW(1085) & ChrW(1090) & ChrW(1088)
        Exit Function
    End If
    If Left(lw, 6) = ChrW(1088) & ChrW(1077) & ChrW(1082) & ChrW(1088) & ChrW(1091) & ChrW(1090) Then
        StemWord = ChrW(1088) & ChrW(1077) & ChrW(1082) & ChrW(1088) & ChrW(1091) & ChrW(1090)
        Exit Function
    End If
    If Left(lw, 5) = ChrW(1096) & ChrW(1090) & ChrW(1091) & ChrW(1088) & ChrW(1084) Then
        StemWord = ChrW(1096) & ChrW(1090) & ChrW(1091) & ChrW(1088) & ChrW(1084)
        Exit Function
    End If
    If Left(lw, 6) = ChrW(1090) & ChrW(1072) & ChrW(1085) & ChrW(1082) & ChrW(1086) & ChrW(1074) Then
        StemWord = ChrW(1090) & ChrW(1072) & ChrW(1085) & ChrW(1082) & ChrW(1086) & ChrW(1074)
        Exit Function
    End If
    If Left(lw, 4) = ChrW(1087) & ChrW(1086) & ChrW(1083) & ChrW(1082) Then
        StemWord = ChrW(1087) & ChrW(1086) & ChrW(1083) & ChrW(1082)
        Exit Function
    End If
    If Left(lw, 6) = ChrW(1082) & ChrW(1086) & ChrW(1088) & ChrW(1087) & ChrW(1091) & ChrW(1089) Then
        StemWord = ChrW(1082) & ChrW(1086) & ChrW(1088) & ChrW(1087) & ChrW(1091) & ChrW(1089)
        Exit Function
    End If

    If Len(lw) > 8 And Right(lw, 6) = ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1075) & ChrW(1086) Then
        StemWord = Left(lw, Len(lw) - 6)
        Exit Function
    End If
    If Len(lw) > 7 And Right(lw, 5) = ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1111) Then
        StemWord = Left(lw, Len(lw) - 5)
        Exit Function
    End If
    If Len(lw) > 8 And Right(lw, 6) = ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1084) & ChrW(1091) Then
        StemWord = Left(lw, Len(lw) - 6)
        Exit Function
    End If
    If Len(lw) > 7 And Right(lw, 5) = ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1080) & ChrW(1084) Then
        StemWord = Left(lw, Len(lw) - 5)
        Exit Function
    End If
    If Len(lw) > 7 And Right(lw, 5) = ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1102) Then
        StemWord = Left(lw, Len(lw) - 5)
        Exit Function
    End If
    If Len(lw) > 8 And Right(lw, 6) = ChrW(1094) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1075) & ChrW(1086) Then
        StemWord = Left(lw, Len(lw) - 6)
        Exit Function
    End If
    If Len(lw) > 7 And Right(lw, 5) = ChrW(1094) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1111) Then
        StemWord = Left(lw, Len(lw) - 5)
        Exit Function
    End If
    If Len(lw) > 8 And Right(lw, 6) = ChrW(1094) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1084) & ChrW(1091) Then
        StemWord = Left(lw, Len(lw) - 6)
        Exit Function
    End If
    If Len(lw) > 7 And Right(lw, 5) = ChrW(1094) & ChrW(1100) & ChrW(1082) & ChrW(1080) & ChrW(1084) Then
        StemWord = Left(lw, Len(lw) - 5)
        Exit Function
    End If
    If Len(lw) > 7 And Right(lw, 5) = ChrW(1094) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1102) Then
        StemWord = Left(lw, Len(lw) - 5)
        Exit Function
    End If
    If Len(lw) > 8 And Right(lw, 6) = ChrW(1079) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1075) & ChrW(1086) Then
        StemWord = Left(lw, Len(lw) - 6)
        Exit Function
    End If
    If Len(lw) > 7 And Right(lw, 5) = ChrW(1079) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1111) Then
        StemWord = Left(lw, Len(lw) - 5)
        Exit Function
    End If
    If Len(lw) > 8 And Right(lw, 6) = ChrW(1079) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1084) & ChrW(1091) Then
        StemWord = Left(lw, Len(lw) - 6)
        Exit Function
    End If
    If Len(lw) > 7 And Right(lw, 5) = ChrW(1079) & ChrW(1100) & ChrW(1082) & ChrW(1080) & ChrW(1084) Then
        StemWord = Left(lw, Len(lw) - 5)
        Exit Function
    End If
    If Len(lw) > 7 And Right(lw, 5) = ChrW(1079) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1102) Then
        StemWord = Left(lw, Len(lw) - 5)
        Exit Function
    End If
    If Len(lw) > 6 And Right(lw, 4) = ChrW(1085) & ChrW(1080) & ChrW(1084) & ChrW(1080) Then
        StemWord = Left(lw, Len(lw) - 4)
        Exit Function
    End If
    If Len(lw) > 5 And Right(lw, 3) = ChrW(1080) & ChrW(1084) & ChrW(1080) Then
        StemWord = Left(lw, Len(lw) - 3)
        Exit Function
    End If
    If Len(lw) > 5 And Right(lw, 3) = ChrW(1072) & ChrW(1084) & ChrW(1080) Then
        StemWord = Left(lw, Len(lw) - 3)
        Exit Function
    End If
    If Len(lw) > 5 And Right(lw, 3) = ChrW(1103) & ChrW(1084) & ChrW(1080) Then
        StemWord = Left(lw, Len(lw) - 3)
        Exit Function
    End If
    If Len(lw) > 5 And Right(lw, 3) = ChrW(1086) & ChrW(1075) & ChrW(1086) Then
        StemWord = Left(lw, Len(lw) - 3)
        Exit Function
    End If
    If Len(lw) > 5 And Right(lw, 3) = ChrW(1086) & ChrW(1084) & ChrW(1091) Then
        StemWord = Left(lw, Len(lw) - 3)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1110) & ChrW(1081) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1080) & ChrW(1081) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1086) & ChrW(1111) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1086) & ChrW(1102) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1080) & ChrW(1093) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1103) & ChrW(1084) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1103) & ChrW(1093) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1072) & ChrW(1093) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1086) & ChrW(1084) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1077) & ChrW(1084) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1108) & ChrW(1084) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1110) & ChrW(1074) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1077) & ChrW(1074) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1108) & ChrW(1074) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1080) & ChrW(1084) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 4 And Right(lw, 2) = ChrW(1110) & ChrW(1084) Then
        StemWord = Left(lw, Len(lw) - 2)
        Exit Function
    End If
    If Len(lw) > 3 And Right(lw, 1) = ChrW(1080) Then
        StemWord = Left(lw, Len(lw) - 1)
        Exit Function
    End If
    If Len(lw) > 3 And Right(lw, 1) = ChrW(1110) Then
        StemWord = Left(lw, Len(lw) - 1)
        Exit Function
    End If
    If Len(lw) > 3 And Right(lw, 1) = ChrW(1072) Then
        StemWord = Left(lw, Len(lw) - 1)
        Exit Function
    End If
    If Len(lw) > 3 And Right(lw, 1) = ChrW(1103) Then
        StemWord = Left(lw, Len(lw) - 1)
        Exit Function
    End If
    If Len(lw) > 3 And Right(lw, 1) = ChrW(1091) Then
        StemWord = Left(lw, Len(lw) - 1)
        Exit Function
    End If
    If Len(lw) > 3 And Right(lw, 1) = ChrW(1102) Then
        StemWord = Left(lw, Len(lw) - 1)
        Exit Function
    End If
    If Len(lw) > 3 And Right(lw, 1) = ChrW(1077) Then
        StemWord = Left(lw, Len(lw) - 1)
        Exit Function
    End If
    If Len(lw) > 3 And Right(lw, 1) = ChrW(1108) Then
        StemWord = Left(lw, Len(lw) - 1)
        Exit Function
    End If
    If Len(lw) > 3 And Right(lw, 1) = ChrW(1086) Then
        StemWord = Left(lw, Len(lw) - 1)
        Exit Function
    End If
    StemWord = lw
End Function

Private Function BuildSearchKeys(ByVal openName As String) As String()
    Dim words() As String, result() As String
    Dim i As Long, cnt As Long, w As String, lw As String
    Dim stopW As String
    stopW = "|" & ChrW(1090) & ChrW(1072) & "|" & ChrW(1110) & "|" & ChrW(1081) & "|" & ChrW(1079) & "|" & ChrW(1079) & ChrW(1110) & "|" & ChrW(1110) & ChrW(1079) & "|" & ChrW(1085) & ChrW(1072) & "|" & ChrW(1074) & "|" & ChrW(1091) & "|" & ChrW(1076) & ChrW(1086) & "|" & ChrW(1074) & ChrW(1110) & ChrW(1076) & "|" & ChrW(1087) & ChrW(1086) & "|" & ChrW(1087) & ChrW(1088) & ChrW(1080) & "|" & ChrW(1079) & ChrW(1072) & "|" & ChrW(1072) & "|" & ChrW(1110) & ChrW(1084) & "|" & ChrW(1110) & ChrW(1084) & ChrW(1077) & ChrW(1085) & ChrW(1110) & "|"

    words = Split(Replace(Replace(Replace(openName, "-", " "), Chr(39), " "), ChrW(8217), " "))
    cnt = 0
    ReDim result(0 To UBound(words))
    For i = 0 To UBound(words)
        w = Trim(words(i))
        If w = "" Then GoTo NextW
        lw = LCase(w)
        If InStr(1, stopW, "|" & lw & "|", vbTextCompare) > 0 Then GoTo NextW
        If Len(w) < 2 Then GoTo NextW
        result(cnt) = StemWord(w)
        cnt = cnt + 1
NextW:
    Next i
    If cnt = 0 Then
        ReDim result(0 To 0)
        result(0) = LCase(openName)
    Else
        ReDim Preserve result(0 To cnt - 1)
    End If
    BuildSearchKeys = result
End Function

Private Function FuzzyMatch(ByVal txt As String, keys() As String) As Boolean
    Dim i As Long, ltxt As String
    ltxt = LCase(txt)
    FuzzyMatch = True
    For i = 0 To UBound(keys)
        If keys(i) <> "" Then
            If InStr(1, ltxt, keys(i), vbTextCompare) = 0 Then
                FuzzyMatch = False
                Exit Function
            End If
        End If
    Next i
End Function

Private Function FindSpan(ByVal pText As String, keys() As String, ByRef outStart As Long, ByRef outEnd As Long) As Boolean
    Dim i As Long, currentPos As Long, matchPos As Long
    Dim ltext As String
    ltext = LCase(pText)
    currentPos = 1
    outStart = -1
    For i = 0 To UBound(keys)
        If keys(i) <> "" Then
            matchPos = InStr(currentPos, ltext, keys(i), vbTextCompare)
            If matchPos = 0 Then
                FindSpan = False
                Exit Function
            End If
            If outStart = -1 Then outStart = matchPos
            currentPos = matchPos + Len(keys(i))
        End If
    Next i
    Do While currentPos <= Len(ltext)
        Dim ch As String
        ch = Mid(ltext, currentPos, 1)
        If InStr(" .,;:" & Chr(34) & Chr(39) & Chr(13) & Chr(10), ch) > 0 Then
            Exit Do
        End If
        currentPos = currentPos + 1
    Loop
    outEnd = currentPos
    FindSpan = True
End Function

Private Sub ContextReplace(rng As Object, findText As String, replaceText As String)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = 0 ' wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .Execute Replace:=2 ' wdReplaceAll
    End With
End Sub

Private Sub HighlightLeftover(rng As Object, findText As String)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Highlight = True
        .Forward = True
        .Wrap = 0
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        rng.Application.Options.DefaultHighlightColorIndex = 7 ' wdYellow
        .Execute Replace:=2
    End With
End Sub

Sub GenerateClosedAndNotice()
    Dim xlApp As Object, xlWb As Object, xlSheet As Object
    Dim wordApp As Object, srcDoc As Object, noticeTemp As Object, closedTemp As Object
    Dim tempDoc As Object, noticeDoc As Object, closedDoc As Object
    Dim fd As Object, fso As Object, outFolder As String
    Dim excelPath As String, docPath As String, noticePath As String, closedPath As String
    Dim weCreatedWord As Boolean, useActiveDoc As Boolean
    Dim mapKeys() As Variant, mapCiphers() As String, mapCorps() As String, mapDestWhere() As String
    Dim mapCount As Long, mc As Long, lastRow As Long, r As Long, i As Long, j As Long
    Dim openName As String, cipher As String, destWhere As String, corps As String
    Dim pText As String, noticeDict As Object, v As Variant, noticeString As String
    Dim outStart As Long, outEnd As Long, spanRng As Object
    Dim p As Object, pCount As Long
    Dim matchedAny As Boolean
    Dim keys() As String
    Dim kwFound As Boolean

    On Error GoTo ErrorHandler

    Set fd = Application.FileDialog(1): fd.Title = ChrW(1054) & ChrW(1073) & ChrW(1077) & ChrW(1088) & ChrW(1110) & ChrW(1090) & ChrW(1100) & " Excel " & ChrW(1092) & ChrW(1072) & ChrW(1081) & ChrW(1083) & " " & ChrW(1042) & ChrW(1063): fd.Filters.Clear: fd.Filters.Add "Excel", "*.xlsx;*.xls": If fd.Show <> -1 Then Exit Sub
    excelPath = fd.SelectedItems(1)

    useActiveDoc = False
    weCreatedWord = False
    docPath = ""
    If Application.Name = "Microsoft Word" Then
        Set wordApp = Application
        If wordApp.Documents.Count > 0 Then
            If MsgBox(wordApp.ActiveDocument.Name & vbCrLf & ChrW(1042) & ChrW(1080) & ChrW(1082) & ChrW(1086) & ChrW(1088) & ChrW(1080) & ChrW(1089) & ChrW(1090) & ChrW(1072) & ChrW(1090) & ChrW(1080) & " " & ChrW(1103) & ChrW(1082) & " " & ChrW(1053) & ChrW(1040) & ChrW(1050) & ChrW(1040) & ChrW(1047) & "?", vbYesNo) = vbYes Then
                Set srcDoc = wordApp.ActiveDocument
                docPath = srcDoc.FullName
                useActiveDoc = True
            End If
        End If
    End If
    If docPath = "" Then
        Set fd = Application.FileDialog(1): fd.Title = ChrW(1054) & ChrW(1073) & ChrW(1077) & ChrW(1088) & ChrW(1110) & ChrW(1090) & ChrW(1100) & " " & ChrW(1092) & ChrW(1072) & ChrW(1081) & ChrW(1083) & " " & ChrW(1053) & ChrW(1040) & ChrW(1050) & ChrW(1040) & ChrW(1047) & ChrW(1059): fd.Filters.Clear: fd.Filters.Add "Word", "*.docx": If fd.Show <> -1 Then Exit Sub
        docPath = fd.SelectedItems(1)
    End If

    Set fd = Application.FileDialog(1): fd.Title = ChrW(1054) & ChrW(1073) & ChrW(1077) & ChrW(1088) & ChrW(1110) & ChrW(1090) & ChrW(1100) & " " & ChrW(1047) & ChrW(1056) & ChrW(1040) & ChrW(1047) & ChrW(1054) & ChrW(1050) & " " & ChrW(1055) & ChrW(1086) & ChrW(1074) & ChrW(1110) & ChrW(1076) & ChrW(1086) & ChrW(1084) & ChrW(1083) & ChrW(1077) & ChrW(1085) & ChrW(1085) & ChrW(1103) & " (" & ChrW(1076) & ChrW(1077) & " " & ChrW(1108) & " {{" & ChrW(1088) & ChrW(1086) & ChrW(1079) & ChrW(1089) & ChrW(1080) & ChrW(1083) & ChrW(1082) & ChrW(1072) & "}})": fd.Filters.Clear: fd.Filters.Add "Word", "*.docx": If fd.Show <> -1 Then Exit Sub
    noticePath = fd.SelectedItems(1)

    Set fd = Application.FileDialog(1): fd.Title = ChrW(1054) & ChrW(1073) & ChrW(1077) & ChrW(1088) & ChrW(1110) & ChrW(1090) & ChrW(1100) & " " & ChrW(1047) & ChrW(1056) & ChrW(1040) & ChrW(1047) & ChrW(1054) & ChrW(1050) & " " & ChrW(1047) & ChrW(1072) & ChrW(1082) & ChrW(1088) & ChrW(1080) & ChrW(1090) & ChrW(1086) & ChrW(1075) & ChrW(1086) & " " & ChrW(1053) & ChrW(1072) & ChrW(1082) & ChrW(1072) & ChrW(1079) & ChrW(1091) & " (" & ChrW(1076) & ChrW(1077) & " " & ChrW(1108) & " {{" & ChrW(1079) & ChrW(1084) & ChrW(1110) & ChrW(1089) & ChrW(1090) & "}})": fd.Filters.Clear: fd.Filters.Add "Word", "*.docx": If fd.Show <> -1 Then Exit Sub
    closedPath = fd.SelectedItems(1)

    Set fso = CreateObject("Scripting.FileSystemObject")
    outFolder = fso.GetParentFolderName(docPath) & "\" & ChrW(1047) & ChrW(1072) & ChrW(1082) & ChrW(1088) & ChrW(1080) & ChrW(1090) & ChrW(1110) & "_" & ChrW(1090) & ChrW(1072) & "_" & ChrW(1055) & ChrW(1086) & ChrW(1074) & ChrW(1110) & ChrW(1076) & ChrW(1086) & ChrW(1084) & ChrW(1083) & ChrW(1077) & ChrW(1085) & ChrW(1085) & ChrW(1103)
    If Not fso.FolderExists(outFolder) Then fso.CreateFolder outFolder

    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo ErrorHandler
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlWb = xlApp.Workbooks.Open(excelPath, ReadOnly:=True)
    Set xlSheet = xlWb.Sheets(1)
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row
    mapCount = lastRow: mc = 0
    ReDim mapKeys(1 To mapCount), mapCiphers(1 To mapCount), mapCorps(1 To mapCount), mapDestWhere(1 To mapCount)
    For r = 2 To lastRow
        openName = Trim(CStr(xlSheet.Cells(r, 1).Value))
        If openName <> "" Then
            mc = mc + 1
            cipher = Trim(CStr(xlSheet.Cells(r, 2).Value))
            If cipher = "" Then cipher = openName
            mapCiphers(mc) = cipher
            mapCorps(mc) = Trim(CStr(xlSheet.Cells(r, 4).Value)) ' Column D
            mapDestWhere(mc) = Trim(CStr(xlSheet.Cells(r, 6).Value)) ' Column F
            mapKeys(mc) = BuildSearchKeys(openName)
        End If
    Next r
    xlWb.Close False
    Set xlApp = Nothing

    If Not useActiveDoc Then
        If wordApp Is Nothing Then Set wordApp = CreateObject("Word.Application"): weCreatedWord = True
        wordApp.Visible = False
        Set srcDoc = wordApp.Documents.Open(docPath, ReadOnly:=True)
    End If

    Set noticeDict = CreateObject("Scripting.Dictionary")
    For i = 1 To srcDoc.Paragraphs.Count
        pText = srcDoc.Paragraphs(i).Range.Text
        For j = 1 To mc
            keys = mapKeys(j)
            If FuzzyMatch(pText, keys) Then
                If mapDestWhere(j) <> "" Then noticeDict(mapDestWhere(j)) = 1
                If mapCorps(j) <> "" Then noticeDict(mapCorps(j)) = 1
                Exit For
            End If
        Next j
    Next i
    noticeString = ""
    For Each v In noticeDict.Keys
        noticeString = noticeString & CStr(v) & vbCrLf
    Next v

    Set noticeTemp = wordApp.Documents.Open(noticePath, ReadOnly:=True)
    With noticeTemp.Content.Find
        .Text = "{{" & ChrW(1088) & ChrW(1086) & ChrW(1079) & ChrW(1089) & ChrW(1080) & ChrW(1083) & ChrW(1082) & ChrW(1072) & "}}"
        .Replacement.Text = noticeString
        .Execute Replace:=2
    End With
    noticeTemp.SaveAs2 outFolder & "\" & ChrW(1055) & ChrW(1086) & ChrW(1074) & ChrW(1110) & ChrW(1076) & ChrW(1086) & ChrW(1084) & ChrW(1083) & ChrW(1077) & ChrW(1085) & ChrW(1085) & ChrW(1103) & "_" & ChrW(1056) & ChrW(1086) & ChrW(1079) & ChrW(1089) & ChrW(1080) & ChrW(1083) & ChrW(1082) & ChrW(1072) & ".docx", 16
    noticeTemp.Close False

    srcDoc.Content.Copy
    Set tempDoc = wordApp.Documents.Add
    tempDoc.Content.PasteAndFormat 16

    pCount = tempDoc.Paragraphs.Count
    For i = 1 To pCount
        Set p = tempDoc.Paragraphs(i)
        pText = p.Range.Text
        matchedAny = False
        For j = 1 To mc
            keys = mapKeys(j)
            If FuzzyMatch(pText, keys) Then
                matchedAny = True
                If FindSpan(pText, keys, outStart, outEnd) Then
                    Set spanRng = p.Range.Duplicate
                    spanRng.Start = spanRng.Start + outStart - 1
                    spanRng.End = spanRng.Start + (outEnd - outStart)
                    spanRng.Text = ChrW(1074) & ChrW(1110) & ChrW(1081) & ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1074) & ChrW(1086) & ChrW(1111) & " " & ChrW(1095) & ChrW(1072) & ChrW(1089) & ChrW(1090) & ChrW(1080) & ChrW(1085) & ChrW(1080) & " " & mapCiphers(j)
                End If
                Exit For
            End If
        Next j

        If Not matchedAny And Len(pText) > 5 Then
            kwFound = False
            If InStr(1, LCase(pText), ChrW(1073) & ChrW(1088) & ChrW(1080) & ChrW(1075) & ChrW(1072) & ChrW(1076), vbTextCompare) > 0 Then kwFound = True
            If InStr(1, LCase(pText), ChrW(1087) & ChrW(1086) & ChrW(1083) & ChrW(1082), vbTextCompare) > 0 Then kwFound = True
            If InStr(1, LCase(pText), ChrW(1073) & ChrW(1072) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1100) & ChrW(1081) & ChrW(1086) & ChrW(1085), vbTextCompare) > 0 Then kwFound = True
            If kwFound Then
                p.Range.HighlightColorIndex = 6 ' wdRed
            End If
        End If

        ContextReplace p.Range, ChrW(1094) & ChrW(1110) & ChrW(1108) & ChrW(1111) & " " & ChrW(1078) & " " & ChrW(1073) & ChrW(1088) & ChrW(1080) & ChrW(1075) & ChrW(1072) & ChrW(1076) & ChrW(1080), ChrW(1094) & ChrW(1110) & ChrW(1108) & ChrW(1111) & " " & ChrW(1078) & " " & ChrW(1074) & ChrW(1110) & ChrW(1081) & ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1074) & ChrW(1086) & ChrW(1111) & " " & ChrW(1095) & ChrW(1072) & ChrW(1089) & ChrW(1090) & ChrW(1080) & ChrW(1085) & ChrW(1080)
        ContextReplace p.Range, ChrW(1094) & ChrW(1100) & ChrW(1086) & ChrW(1075) & ChrW(1086) & " " & ChrW(1078) & " " & ChrW(1087) & ChrW(1086) & ChrW(1083) & ChrW(1082) & ChrW(1091), ChrW(1094) & ChrW(1110) & ChrW(1108) & ChrW(1111) & " " & ChrW(1078) & " " & ChrW(1074) & ChrW(1110) & ChrW(1081) & ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1074) & ChrW(1086) & ChrW(1111) & " " & ChrW(1095) & ChrW(1072) & ChrW(1089) & ChrW(1090) & ChrW(1080) & ChrW(1085) & ChrW(1080)
        ContextReplace p.Range, ChrW(1094) & ChrW(1100) & ChrW(1086) & ChrW(1075) & ChrW(1086) & " " & ChrW(1078) & " " & ChrW(1073) & ChrW(1072) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1100) & ChrW(1081) & ChrW(1086) & ChrW(1085) & ChrW(1091), ChrW(1094) & ChrW(1110) & ChrW(1108) & ChrW(1111) & " " & ChrW(1078) & " " & ChrW(1074) & ChrW(1110) & ChrW(1081) & ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1074) & ChrW(1086) & ChrW(1111) & " " & ChrW(1095) & ChrW(1072) & ChrW(1089) & ChrW(1090) & ChrW(1080) & ChrW(1085) & ChrW(1080)
        ContextReplace p.Range, ChrW(1094) & ChrW(1100) & ChrW(1086) & ChrW(1075) & ChrW(1086) & " " & ChrW(1078) & " " & ChrW(1094) & ChrW(1077) & ChrW(1085) & ChrW(1090) & ChrW(1088) & ChrW(1091), ChrW(1094) & ChrW(1110) & ChrW(1108) & ChrW(1111) & " " & ChrW(1078) & " " & ChrW(1074) & ChrW(1110) & ChrW(1081) & ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1074) & ChrW(1086) & ChrW(1111) & " " & ChrW(1095) & ChrW(1072) & ChrW(1089) & ChrW(1090) & ChrW(1080) & ChrW(1085) & ChrW(1080)
        ContextReplace p.Range, ChrW(1079) & ChrW(1072) & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1077) & ChrW(1085) & ChrW(1086) & ChrW(1111) & " " & ChrW(1073) & ChrW(1088) & ChrW(1080) & ChrW(1075) & ChrW(1072) & ChrW(1076) & ChrW(1080), ChrW(1079) & ChrW(1072) & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1077) & ChrW(1085) & ChrW(1086) & ChrW(1111) & " " & ChrW(1074) & ChrW(1110) & ChrW(1081) & ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1074) & ChrW(1086) & ChrW(1111) & " " & ChrW(1095) & ChrW(1072) & ChrW(1089) & ChrW(1090) & ChrW(1080) & ChrW(1085) & ChrW(1080)
        ContextReplace p.Range, ChrW(1079) & ChrW(1072) & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1077) & ChrW(1085) & ChrW(1086) & ChrW(1075) & ChrW(1086) & " " & ChrW(1087) & ChrW(1086) & ChrW(1083) & ChrW(1082) & ChrW(1091), ChrW(1079) & ChrW(1072) & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1077) & ChrW(1085) & ChrW(1086) & ChrW(1111) & " " & ChrW(1074) & ChrW(1110) & ChrW(1081) & ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1074) & ChrW(1086) & ChrW(1111) & " " & ChrW(1095) & ChrW(1072) & ChrW(1089) & ChrW(1090) & ChrW(1080) & ChrW(1085) & ChrW(1080)
        ContextReplace p.Range, ChrW(1079) & ChrW(1072) & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1077) & ChrW(1085) & ChrW(1086) & ChrW(1075) & ChrW(1086) & " " & ChrW(1073) & ChrW(1072) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1100) & ChrW(1081) & ChrW(1086) & ChrW(1085) & ChrW(1091), ChrW(1079) & ChrW(1072) & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1077) & ChrW(1085) & ChrW(1086) & ChrW(1111) & " " & ChrW(1074) & ChrW(1110) & ChrW(1081) & ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1074) & ChrW(1086) & ChrW(1111) & " " & ChrW(1095) & ChrW(1072) & ChrW(1089) & ChrW(1090) & ChrW(1080) & ChrW(1085) & ChrW(1080)
        ContextReplace p.Range, ChrW(1094) & ChrW(1110) & ChrW(1108) & ChrW(1111) & " " & ChrW(1089) & ChrW(1072) & ChrW(1084) & ChrW(1086) & ChrW(1111) & " " & ChrW(1073) & ChrW(1088) & ChrW(1080) & ChrW(1075) & ChrW(1072) & ChrW(1076) & ChrW(1080), ChrW(1094) & ChrW(1110) & ChrW(1108) & ChrW(1111) & " " & ChrW(1078) & " " & ChrW(1074) & ChrW(1110) & ChrW(1081) & ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1074) & ChrW(1086) & ChrW(1111) & " " & ChrW(1095) & ChrW(1072) & ChrW(1089) & ChrW(1090) & ChrW(1080) & ChrW(1085) & ChrW(1080)
        ContextReplace p.Range, ChrW(1094) & ChrW(1100) & ChrW(1086) & ChrW(1075) & ChrW(1086) & " " & ChrW(1089) & ChrW(1072) & ChrW(1084) & ChrW(1086) & ChrW(1075) & ChrW(1086) & " " & ChrW(1087) & ChrW(1086) & ChrW(1083) & ChrW(1082) & ChrW(1091), ChrW(1094) & ChrW(1110) & ChrW(1108) & ChrW(1111) & " " & ChrW(1078) & " " & ChrW(1074) & ChrW(1110) & ChrW(1081) & ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1074) & ChrW(1086) & ChrW(1111) & " " & ChrW(1095) & ChrW(1072) & ChrW(1089) & ChrW(1090) & ChrW(1080) & ChrW(1085) & ChrW(1080)
        ContextReplace p.Range, ChrW(1094) & ChrW(1100) & ChrW(1086) & ChrW(1075) & ChrW(1086) & " " & ChrW(1089) & ChrW(1072) & ChrW(1084) & ChrW(1086) & ChrW(1075) & ChrW(1086) & " " & ChrW(1073) & ChrW(1072) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1100) & ChrW(1081) & ChrW(1086) & ChrW(1085) & ChrW(1091), ChrW(1094) & ChrW(1110) & ChrW(1108) & ChrW(1111) & " " & ChrW(1078) & " " & ChrW(1074) & ChrW(1110) & ChrW(1081) & ChrW(1089) & ChrW(1100) & ChrW(1082) & ChrW(1086) & ChrW(1074) & ChrW(1086) & ChrW(1111) & " " & ChrW(1095) & ChrW(1072) & ChrW(1089) & ChrW(1090) & ChrW(1080) & ChrW(1085) & ChrW(1080)

        HighlightLeftover p.Range, ChrW(1073) & ChrW(1088) & ChrW(1080) & ChrW(1075) & ChrW(1072) & ChrW(1076)
        HighlightLeftover p.Range, ChrW(1087) & ChrW(1086) & ChrW(1083) & ChrW(1082)
        HighlightLeftover p.Range, ChrW(1073) & ChrW(1072) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1100) & ChrW(1081) & ChrW(1086) & ChrW(1085)
    Next i

    Set closedTemp = wordApp.Documents.Open(closedPath, ReadOnly:=True)
    With closedTemp.Content.Find
        .Text = "{{" & ChrW(1079) & ChrW(1084) & ChrW(1110) & ChrW(1089) & ChrW(1090) & "}}"
        .MatchWholeWord = False
        .MatchCase = False
        If .Execute Then
            .Parent.Text = ""
            .Parent.Collapse 1 ' Start
            tempDoc.Content.Copy
            .Parent.PasteAndFormat 16
        End If
    End With
    closedTemp.SaveAs2 outFolder & "\" & ChrW(1047) & ChrW(1072) & ChrW(1082) & ChrW(1088) & ChrW(1080) & ChrW(1090) & ChrW(1080) & ChrW(1081) & "_" & ChrW(1053) & ChrW(1072) & ChrW(1082) & ChrW(1072) & ChrW(1079) & ".docx", 16
    closedTemp.Close False
    tempDoc.Close False

    If Not useActiveDoc Then srcDoc.Close False
    If weCreatedWord Then wordApp.Quit
    Set wordApp = Nothing
    MsgBox ChrW(1043) & ChrW(1086) & ChrW(1090) & ChrW(1086) & ChrW(1074) & ChrW(1086) & "! " & ChrW(1047) & ChrW(1075) & ChrW(1077) & ChrW(1085) & ChrW(1077) & ChrW(1088) & ChrW(1086) & ChrW(1074) & ChrW(1072) & ChrW(1085) & ChrW(1086) & " " & ChrW(1055) & ChrW(1086) & ChrW(1074) & ChrW(1110) & ChrW(1076) & ChrW(1086) & ChrW(1084) & ChrW(1083) & ChrW(1077) & ChrW(1085) & ChrW(1085) & ChrW(1103) & " " & ChrW(1090) & ChrW(1072) & " " & ChrW(1047) & ChrW(1072) & ChrW(1082) & ChrW(1088) & ChrW(1080) & ChrW(1090) & ChrW(1080) & ChrW(1081) & " " & ChrW(1053) & ChrW(1072) & ChrW(1082) & ChrW(1072) & ChrW(1079) & " " & ChrW(1091) & " " & ChrW(1087) & ChrW(1072) & ChrW(1087) & ChrW(1082) & ChrW(1091) & ": " & vbCrLf & outFolder, vbInformation
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not noticeTemp Is Nothing Then noticeTemp.Close False
    If Not closedTemp Is Nothing Then closedTemp.Close False
    If Not tempDoc Is Nothing Then tempDoc.Close False
    If Not useActiveDoc And Not srcDoc Is Nothing Then srcDoc.Close False
    If weCreatedWord And Not wordApp Is Nothing Then wordApp.Quit
    On Error GoTo 0
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub