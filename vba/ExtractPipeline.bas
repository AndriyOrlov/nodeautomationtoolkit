Attribute VB_Name = "ExtractPipeline"
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

    Dim endLen As Long
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
    ' Returns array of stemmed anchor words (skip stop words, keep numbers)
    Dim words() As String, result() As String
    Dim i As Long, cnt As Long, w As String, lw As String
    Dim stopW As String
    stopW = "|" & ChrW(1090) & ChrW(1072) & "|" & ChrW(1110) & "|" & ChrW(1081) & "|" & ChrW(1079) & "|" & ChrW(1079) & ChrW(1110) & "|" & ChrW(1110) & ChrW(1079) & "|" & ChrW(1085) & ChrW(1072) & "|" & ChrW(1074) & "|" & ChrW(1091) & "|" & ChrW(1076) & ChrW(1086) & "|" & ChrW(1074) & ChrW(1110) & ChrW(1076) & "|" & ChrW(1087) & ChrW(1086) & "|" & ChrW(1087) & ChrW(1088) & ChrW(1080) & "|" & ChrW(1079) & ChrW(1072) & "|" & ChrW(1072) & "|" & ChrW(1110) & ChrW(1084) & "|" & ChrW(1110) & ChrW(1084) & ChrW(1077) & ChrW(1085) & ChrW(1110) & "|"

    words = Split(Replace(Replace(openName, "-", " "), Chr(39), " "))
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

Sub GenerateExtracts()
    Dim xlApp As Object, xlWb As Object, xlSheet As Object
    Dim wordApp As Object, srcDoc As Object, targetDoc As Object
    Dim fso As Object, fd As Object, rng As Object
    Dim excelPath As String, docPath As String, outFolder As String
    Dim weCreatedWord As Boolean, xlWasRunning As Boolean, useActiveDoc As Boolean
    Dim mapKeys() As Variant
    Dim mapCiphers() As String
    Dim mapRecipTo() As String
    Dim mapDestWhere() As String
    Dim mapCount As Long
    Dim itemParaMap As Object
    Dim lastRow As Long, r As Long, i As Long, j As Long, idx As Long
    Dim openName As String, cipher As String, recipTo As String, destWhere As String, abbr As String
    Dim pText As String
    Dim col As Collection
    Dim k As Variant, paraIdx As Variant
    Dim reportMsg As String, totalItems As Long
    Dim cleanCipher As String, outFile As String, pIndex As Long
    Dim ans As VbMsgBoxResult
    Dim matched As Boolean
    Dim globalFirstPara As Long, globalLastPara As Long
    Dim unitIndex As Long

    On Error GoTo ErrorHandler

    Set fd = Application.FileDialog(1)
    fd.Title = ChrW(1054) & ChrW(1073) & ChrW(1077) & ChrW(1088) & ChrW(1110) & ChrW(1090) & ChrW(1100) & " Excel " & ChrW(1092) & ChrW(1072) & ChrW(1081) & ChrW(1083) & " " & ChrW(1074) & ChrW(1110) & ChrW(1076) & ChrW(1087) & ChrW(1086) & ChrW(1074) & ChrW(1110) & ChrW(1076) & ChrW(1085) & ChrW(1086) & ChrW(1089) & ChrW(1090) & ChrW(1077) & ChrW(1081) & " " & ChrW(1042) & ChrW(1063) & " (" & ChrW(1082) & ChrW(1086) & ChrW(1083) & ChrW(1086) & ChrW(1085) & ChrW(1082) & ChrW(1080) & " A-F)"
    fd.Filters.Clear
    fd.Filters.Add "Excel", "*.xlsx;*.xls;*.xlsm"
    If fd.Show <> -1 Then Exit Sub
    excelPath = fd.SelectedItems(1)

    useActiveDoc = False
    weCreatedWord = False
    docPath = ""
    If Application.Name = "Microsoft Word" Then
        Set wordApp = Application
        If wordApp.Documents.Count > 0 Then
            ans = MsgBox(wordApp.ActiveDocument.Name & vbCrLf & vbCrLf & ChrW(1042) & ChrW(1080) & ChrW(1082) & ChrW(1086) & ChrW(1088) & ChrW(1080) & ChrW(1089) & ChrW(1090) & ChrW(1072) & ChrW(1090) & ChrW(1080) & " " & ChrW(1074) & ChrW(1110) & ChrW(1076) & ChrW(1082) & ChrW(1088) & ChrW(1080) & ChrW(1090) & ChrW(1080) & ChrW(1081) & " " & ChrW(1076) & ChrW(1086) & ChrW(1082) & ChrW(1091) & ChrW(1084) & ChrW(1077) & ChrW(1085) & ChrW(1090) & " " & ChrW(1103) & ChrW(1082) & " " & ChrW(1085) & ChrW(1072) & ChrW(1082) & ChrW(1072) & ChrW(1079) & "?", vbQuestion + vbYesNo, ChrW(1040) & ChrW(1082) & ChrW(1090) & ChrW(1080) & ChrW(1074) & ChrW(1085) & ChrW(1080) & ChrW(1081) & " " & ChrW(1076) & ChrW(1086) & ChrW(1082) & ChrW(1091) & ChrW(1084) & ChrW(1077) & ChrW(1085) & ChrW(1090))
            If ans = vbYes Then
                Set srcDoc = wordApp.ActiveDocument
                docPath = srcDoc.FullName
                useActiveDoc = True
            End If
        End If
    End If
    If docPath = "" Then
        Set fd = Application.FileDialog(1)
        fd.Title = ChrW(1054) & ChrW(1073) & ChrW(1077) & ChrW(1088) & ChrW(1110) & ChrW(1090) & ChrW(1100) & " " & ChrW(1092) & ChrW(1072) & ChrW(1081) & ChrW(1083) & " " & ChrW(1086) & ChrW(1088) & ChrW(1080) & ChrW(1075) & ChrW(1110) & ChrW(1085) & ChrW(1072) & ChrW(1083) & ChrW(1100) & ChrW(1085) & ChrW(1086) & ChrW(1075) & ChrW(1086) & " " & ChrW(1085) & ChrW(1072) & ChrW(1082) & ChrW(1072) & ChrW(1079) & ChrW(1091) & " (.docx)"
        fd.Filters.Clear
        fd.Filters.Add "Word", "*.docx;*.doc"
        If fd.Show <> -1 Then Exit Sub
        docPath = fd.SelectedItems(1)
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    outFolder = fso.GetParentFolderName(docPath) & "\Extracts_VBA"
    If Not fso.FolderExists(outFolder) Then fso.CreateFolder outFolder

    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    xlWasRunning = Not (xlApp Is Nothing)
    On Error GoTo ErrorHandler
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlWb = xlApp.Workbooks.Open(excelPath, ReadOnly:=True)
    Set xlSheet = xlWb.Sheets(1)

    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row
    mapCount = lastRow - 1
    If mapCount < 1 Then mapCount = 1
    ReDim mapKeys(1 To mapCount)
    ReDim mapCiphers(1 To mapCount)
    ReDim mapRecipTo(1 To mapCount)
    ReDim mapDestWhere(1 To mapCount)

    Dim mc As Long
    mc = 0
    For r = 2 To lastRow
        openName = Trim(CStr(xlSheet.Cells(r, 1).Value))
        cipher = Trim(CStr(xlSheet.Cells(r, 2).Value))
        abbr = Trim(CStr(xlSheet.Cells(r, 3).Value))
        recipTo = Trim(CStr(xlSheet.Cells(r, 5).Value))
        destWhere = Trim(CStr(xlSheet.Cells(r, 6).Value))
        If openName <> "" Then
            If cipher = "" Then cipher = openName
            mc = mc + 1
            Dim searchStr As String
            searchStr = openName
            If abbr <> "" And abbr <> openName Then searchStr = searchStr & " " & abbr
            mapKeys(mc) = BuildSearchKeys(openName)
            mapCiphers(mc) = cipher
            mapRecipTo(mc) = recipTo
            mapDestWhere(mc) = destWhere
        End If
    Next r
    xlWb.Close False
    If Not xlWasRunning Then xlApp.Quit
    Set xlApp = Nothing

    If mc = 0 Then
        MsgBox "Excel " & ChrW(1087) & ChrW(1086) & ChrW(1088) & ChrW(1086) & ChrW(1078) & ChrW(1085) & ChrW(1110) & ChrW(1081) & " " & ChrW(1072) & ChrW(1073) & ChrW(1086) & " " & ChrW(1085) & ChrW(1077) & " " & ChrW(1084) & ChrW(1110) & ChrW(1089) & ChrW(1090) & ChrW(1080) & ChrW(1090) & ChrW(1100) & " " & ChrW(1076) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1093) & "!", vbExclamation
        Exit Sub
    End If

    If Not useActiveDoc Then
        On Error Resume Next
        Set wordApp = GetObject(, "Word.Application")
        On Error GoTo ErrorHandler
        If wordApp Is Nothing Then
            Set wordApp = CreateObject("Word.Application")
            weCreatedWord = True
        End If
        wordApp.Visible = False
        wordApp.DisplayAlerts = False
        Set srcDoc = wordApp.Documents.Open(docPath, ReadOnly:=True)
    End If

    globalFirstPara = srcDoc.Paragraphs.Count + 1
    globalLastPara = 0
    Set itemParaMap = CreateObject("Scripting.Dictionary")
    For i = 1 To srcDoc.Paragraphs.Count
        pText = srcDoc.Paragraphs(i).Range.Text
        For j = 1 To mc
            Dim keys() As String
            keys = mapKeys(j)
            If FuzzyMatch(pText, keys) Then
                cipher = mapCiphers(j)
                If Not itemParaMap.Exists(cipher) Then
                    Set col = New Collection
                    itemParaMap.Add cipher, col
                End If
                itemParaMap(cipher).Add i
                If i < globalFirstPara Then globalFirstPara = i
                If i > globalLastPara Then globalLastPara = i
                Exit For
            End If
        Next j
    Next i

    If itemParaMap.Count = 0 Then
        MsgBox ChrW(1047) & ChrW(1073) & ChrW(1110) & ChrW(1075) & ChrW(1110) & ChrW(1074) & " " & ChrW(1085) & ChrW(1077) & " " & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1081) & ChrW(1076) & ChrW(1077) & ChrW(1085) & ChrW(1086) & "! " & ChrW(1055) & ChrW(1077) & ChrW(1088) & ChrW(1077) & ChrW(1074) & ChrW(1110) & ChrW(1088) & ChrW(1090) & ChrW(1077) & " " & ChrW(1097) & ChrW(1086) & " " & ChrW(1085) & ChrW(1072) & ChrW(1079) & ChrW(1074) & ChrW(1080) & " " & ChrW(1042) & ChrW(1063) & " " & ChrW(1074) & " Excel " & ChrW(1079) & ChrW(1073) & ChrW(1110) & ChrW(1075) & ChrW(1072) & ChrW(1102) & ChrW(1090) & ChrW(1100) & ChrW(1089) & ChrW(1103) & " " & ChrW(1079) & " " & ChrW(1090) & ChrW(1077) & ChrW(1082) & ChrW(1089) & ChrW(1090) & ChrW(1086) & ChrW(1084) & " " & ChrW(1085) & ChrW(1072) & ChrW(1082) & ChrW(1072) & ChrW(1079) & ChrW(1091) & ".", vbExclamation, ChrW(1055) & ChrW(1086) & ChrW(1084) & ChrW(1080) & ChrW(1083) & ChrW(1082) & ChrW(1072)
        If Not useActiveDoc And Not srcDoc Is Nothing Then srcDoc.Close False
        If weCreatedWord Then wordApp.Quit
        Exit Sub
    End If

    reportMsg = "=== " & ChrW(1047) & ChrW(1042) & ChrW(1030) & ChrW(1058) & " " & ChrW(1056) & ChrW(1054) & ChrW(1047) & ChrW(1055) & ChrW(1054) & ChrW(1044) & ChrW(1030) & ChrW(1051) & ChrW(1059) & " " & ChrW(1055) & ChrW(1059) & ChrW(1053) & ChrW(1050) & ChrW(1058) & ChrW(1030) & ChrW(1042) & " ===" & vbCrLf & vbCrLf
    totalItems = 0
    For Each k In itemParaMap.Keys
        cipher = CStr(k)
        Set col = itemParaMap(cipher)
        recipTo = ""
        For j = 1 To mc
            If mapCiphers(j) = cipher Then
                recipTo = mapRecipTo(j)
                Exit For
            End If
        Next j
        reportMsg = reportMsg & "[" & cipher & "]"
        If recipTo <> "" Then reportMsg = reportMsg & " -> " & recipTo
        reportMsg = reportMsg & " (" & col.Count & ")" & vbCrLf
        For idx = 1 To col.Count
            pText = Trim(srcDoc.Paragraphs(CLng(col(idx))).Range.Text)
            If Len(pText) > 70 Then pText = Left(pText, 70) & "..."
            reportMsg = reportMsg & "  - " & pText & vbCrLf
            totalItems = totalItems + 1
        Next idx
        reportMsg = reportMsg & vbCrLf
    Next k

    If Len(reportMsg) > 800 Then
        reportMsg = "=== " & ChrW(1047) & ChrW(1042) & ChrW(1030) & ChrW(1058) & " " & ChrW(1056) & ChrW(1054) & ChrW(1047) & ChrW(1055) & ChrW(1054) & ChrW(1044) & ChrW(1030) & ChrW(1051) & ChrW(1059) & " " & ChrW(1055) & ChrW(1059) & ChrW(1053) & ChrW(1050) & ChrW(1058) & ChrW(1030) & ChrW(1042) & " ===" & vbCrLf & vbCrLf
        For Each k In itemParaMap.Keys
            cipher = CStr(k)
            Set col = itemParaMap(cipher)
            reportMsg = reportMsg & cipher & " (" & col.Count & ")" & vbCrLf
        Next k
    End If
    reportMsg = reportMsg & vbCrLf & ChrW(1042) & ChrW(1089) & ChrW(1100) & ChrW(1086) & ChrW(1075) & ChrW(1086) & ": " & itemParaMap.Count & "/" & totalItems
    reportMsg = reportMsg & vbCrLf & vbCrLf & ChrW(1043) & ChrW(1077) & ChrW(1085) & ChrW(1077) & ChrW(1088) & ChrW(1091) & ChrW(1074) & ChrW(1072) & ChrW(1090) & ChrW(1080) & " .docx " & ChrW(1074) & ChrW(1080) & ChrW(1090) & ChrW(1103) & ChrW(1075) & ChrW(1080) & "?"
    ans = MsgBox(reportMsg, vbQuestion + vbYesNo, ChrW(1047) & ChrW(1074) & ChrW(1110) & ChrW(1090))
    If ans <> vbYes Then
        If Not useActiveDoc Then srcDoc.Close False
        If weCreatedWord Then wordApp.Quit
        Exit Sub
    End If

    outFile = outFolder & "\All_Extracts.docx"
    Set targetDoc = wordApp.Documents.Add
    targetDoc.PageSetup.TopMargin = 56.7
    targetDoc.PageSetup.BottomMargin = 56.7
    targetDoc.PageSetup.LeftMargin = 56.7
    targetDoc.PageSetup.RightMargin = 42.55

    unitIndex = 0
    For Each k In itemParaMap.Keys
        cipher = CStr(k)
        Set col = itemParaMap(cipher)
        recipTo = ""
        destWhere = ""
        For j = 1 To mc
            If mapCiphers(j) = cipher Then
                recipTo = mapRecipTo(j)
                destWhere = mapDestWhere(j)
                Exit For
            End If
        Next j

        If unitIndex > 0 Then
            Set rng = targetDoc.Content: rng.Collapse 0
            rng.InsertBreak 7 ' wdPageBreak
        End If
        unitIndex = unitIndex + 1

        If recipTo <> "" Then
            Set rng = targetDoc.Content: rng.Collapse 0
            rng.Text = recipTo & vbCr
            rng.Font.Name = "Times New Roman": rng.Font.Size = 14: rng.Font.Bold = True: rng.Font.Italic = False
            rng.ParagraphFormat.Alignment = 2
        End If
        If destWhere <> "" Then
            Set rng = targetDoc.Content: rng.Collapse 0
            rng.Text = destWhere & vbCr
            rng.Font.Name = "Times New Roman": rng.Font.Size = 12: rng.Font.Italic = True: rng.Font.Bold = False
            rng.ParagraphFormat.Alignment = 2
        End If
        Set rng = targetDoc.Content: rng.Collapse 0
        rng.Text = vbCr & ChrW(1042) & ChrW(1048) & ChrW(1058) & ChrW(1071) & ChrW(1043) & " " & ChrW(1047) & " " & ChrW(1053) & ChrW(1040) & ChrW(1050) & ChrW(1040) & ChrW(1047) & ChrW(1059) & vbCr
        rng.Font.Name = "Times New Roman": rng.Font.Size = 16: rng.Font.Bold = True: rng.Font.Italic = False
        rng.ParagraphFormat.Alignment = 1

        If globalFirstPara > 1 And globalFirstPara <= srcDoc.Paragraphs.Count Then
            srcDoc.Range(srcDoc.Paragraphs(1).Range.Start, srcDoc.Paragraphs(globalFirstPara - 1).Range.End).Copy
            Set rng = targetDoc.Content: rng.Collapse 0: rng.PasteAndFormat 16 ' wdFormatOriginalFormatting
        End If

        For Each paraIdx In col
            pIndex = CLng(paraIdx)
            If pIndex >= 1 And pIndex <= srcDoc.Paragraphs.Count Then
                srcDoc.Paragraphs(pIndex).Range.Copy
                Set rng = targetDoc.Content: rng.Collapse 0: rng.PasteAndFormat 16 ' wdFormatOriginalFormatting
            End If
        Next paraIdx

        If globalLastPara > 0 And globalLastPara < srcDoc.Paragraphs.Count Then
            srcDoc.Range(srcDoc.Paragraphs(globalLastPara + 1).Range.Start, srcDoc.Paragraphs(srcDoc.Paragraphs.Count).Range.End).Copy
            Set rng = targetDoc.Content: rng.Collapse 0: rng.PasteAndFormat 16 ' wdFormatOriginalFormatting
        End If

        Set rng = targetDoc.Content: rng.Collapse 0
        rng.Text = vbCr & vbCr & ChrW(1047) & ChrW(1075) & ChrW(1110) & ChrW(1076) & ChrW(1085) & ChrW(1086) & " " & ChrW(1079) & " " & ChrW(1086) & ChrW(1088) & ChrW(1080) & ChrW(1075) & ChrW(1110) & ChrW(1085) & ChrW(1072) & ChrW(1083) & ChrW(1086) & ChrW(1084) & ":" & vbCr
        rng.Font.Name = "Times New Roman": rng.Font.Size = 14: rng.Font.Bold = True: rng.Font.Italic = True
        rng.ParagraphFormat.Alignment = 0

    Next k

    targetDoc.SaveAs2 outFile, 16 ' wdFormatDocumentDefault
    targetDoc.Close False

    If Not useActiveDoc Then srcDoc.Close False
    If weCreatedWord Then wordApp.Quit
    Set wordApp = Nothing
    MsgBox ChrW(1059) & ChrW(1089) & ChrW(1087) & ChrW(1110) & ChrW(1093) & "! " & ChrW(1060) & ChrW(1072) & ChrW(1081) & ChrW(1083) & ChrW(1110) & ChrW(1074) & ": " & itemParaMap.Count & vbCrLf & outFolder, vbInformation, ChrW(1043) & ChrW(1086) & ChrW(1090) & ChrW(1086) & ChrW(1074) & ChrW(1086)
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not useActiveDoc And Not srcDoc Is Nothing Then srcDoc.Close False
    If weCreatedWord And Not wordApp Is Nothing Then wordApp.Quit
    On Error GoTo 0
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub