Attribute VB_Name = "NATProgramCheck"
' NAT - check the open order with the generator program (full check).
'
' Macro: NAT_CheckOrderWithProgram
'   1. takes the text of the active document (with Word auto-numbers);
'   2. runs the generator program in the background:
'        generate_extracts_qt.py --review-text <in> --order-name <name> --out <out>
'      (the same check as the "Check order" button: numbering, RNOKPP, nomenclature,
'       addressees from the table, history of persons and positions from the index);
'   3. marks the document: red = error, yellow = warning, a Word comment on each.
' The document is NOT saved.
'
' Needs the NATCheckCore module (marking, rules) in the same template.
' The program folder is asked once and remembered (HKCU\Software\VB and VBA Program
' Settings\NATCheck). ASCII only: Word imports .bas in the ANSI code page.
Option Explicit

Private Function NAT_ProgramFolder(ByVal ask As Boolean) As String
    Dim folder As String
    folder = GetSetting("NATCheck", "Paths", "Program", "")
    If folder <> "" And Not ask Then
        NAT_ProgramFolder = folder
        Exit Function
    End If
    With Application.FileDialog(4)
        .Title = "NAT: generator program folder"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folder = .SelectedItems(1)
            SaveSetting "NATCheck", "Paths", "Program", folder
        End If
    End With
    NAT_ProgramFolder = folder
End Function

' Command without arguments: built exe, portable python, or the developer .venv.
Private Function NAT_ProgramCommand(ByVal folder As String) As String
    Dim fso As Object, script As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    NAT_ProgramCommand = ""
    If fso.FileExists(folder & "\GeneratorVytyagivQt.exe") Then
        NAT_ProgramCommand = """" & folder & "\GeneratorVytyagivQt.exe"""
        Exit Function
    End If
    script = """" & folder & "\generate_extracts_qt.py"""
    If Not fso.FileExists(folder & "\generate_extracts_qt.py") Then Exit Function
    If fso.FileExists(folder & "\python\python.exe") Then
        NAT_ProgramCommand = """" & folder & "\python\python.exe"" " & script
    ElseIf fso.FileExists(folder & "\.venv\Scripts\python.exe") Then
        NAT_ProgramCommand = """" & folder & "\.venv\Scripts\python.exe"" " & script
    Else
        NAT_ProgramCommand = "python " & script
    End If
End Function

Private Function NAT_ProgramRulesPath(ByVal folder As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(folder & "\word_macro\nat_rules.txt") Then
        NAT_ProgramRulesPath = folder & "\word_macro\nat_rules.txt"
    Else
        NAT_ProgramRulesPath = NAT_RulesPath()
    End If
End Function

Public Sub NAT_CheckOrderWithProgram()
    Dim folder As String, cmd As String, rulesPath As String, rules As Object
    Dim fso As Object, tempFolder As String, inPath As String, outPath As String
    Dim code As Long, text As String, lines As Variant, i As Long, findings As Object, notes As String

    folder = NAT_ProgramFolder(False)
    If folder = "" Then Exit Sub
    cmd = NAT_ProgramCommand(folder)
    If cmd = "" Then
        folder = NAT_ProgramFolder(True)
        If folder = "" Then Exit Sub
        cmd = NAT_ProgramCommand(folder)
    End If
    rulesPath = NAT_ProgramRulesPath(folder)
    If rulesPath = "" Then Exit Sub
    Set rules = NAT_LoadRules(rulesPath)
    If cmd = "" Then
        MsgBox NAT_Msg(rules, "err.program", Array(folder)), vbExclamation, NAT_Msg(rules, "done.title", Array())
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempFolder = fso.GetSpecialFolder(2).Path
    inPath = tempFolder & "\nat_review_in.txt"
    outPath = tempFolder & "\nat_review_out.tsv"
    If fso.FileExists(outPath) Then fso.DeleteFile outPath
    NAT_WriteUtf8 inPath, NAT_DocumentText(ActiveDocument)

    Application.StatusBar = NAT_Msg(rules, "done.title", Array()) & "..."
    code = CreateObject("WScript.Shell").Run(cmd & " --review-text """ & inPath & """ --order-name """ _
        & ActiveDocument.Name & """ --out """ & outPath & """", 0, True)
    Application.StatusBar = ""
    If fso.FileExists(inPath) Then fso.DeleteFile inPath
    If code <> 0 Or Not fso.FileExists(outPath) Then
        MsgBox NAT_Msg(rules, "err.run", Array(code)), vbExclamation, NAT_Msg(rules, "done.title", Array())
        Exit Sub
    End If

    text = NAT_ReadUtf8(outPath)
    fso.DeleteFile outPath
    Set findings = CreateObject("Scripting.Dictionary")
    notes = ""
    lines = Split(Replace(text, vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        If Left(lines(i), 6) = "#note" & vbTab Then
            If notes <> "" Then notes = notes & vbLf
            notes = notes & Mid(lines(i), 7)
        ElseIf lines(i) <> "" And Left(lines(i), 1) <> "#" Then
            findings.Add findings.Count, lines(i)
        End If
    Next
    NAT_ApplyFindings ActiveDocument, findings, rules, notes
End Sub

' Forget the remembered program folder (next run asks again).
Public Sub NAT_ChooseProgramFolder()
    Call NAT_ProgramFolder(True)
End Sub
