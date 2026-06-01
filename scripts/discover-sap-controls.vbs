' Diagnostic: SAP session attach + control tree dump (same attach flow as capture-sap-snapshot.vbs).
' Usage: cscript //Nologo discover-sap-controls.vbs
Option Explicit

Sub LogStep(msg)
    WScript.Echo msg
End Sub

Function IIf(cond, a, b)
    If cond Then IIf = a Else IIf = b
End Function

Sub WalkTree(ctl, depth, ByRef lines, ByRef gridCount, ByRef visitCount)
    On Error Resume Next
    Dim kids, i, child, t, id, rowCount, colCount, indent, line, childCount
    If ctl Is Nothing Or depth > 24 Then Exit Sub

    visitCount = visitCount + 1
    t = ctl.Type
    id = ctl.Id
    rowCount = 0
    colCount = 0
    rowCount = CLng(ctl.RowCount)
    colCount = CLng(ctl.ColumnCount)
    indent = Space(depth * 2)
    line = indent & t
    If Len(id) > 0 Then line = line & " | " & id
    If rowCount > 0 Or colCount > 0 Then
        line = line & " | rows=" & rowCount & " cols=" & colCount
        gridCount = gridCount + 1
    End If
    lines.Add line, True

    Set kids = ctl.Children
    If Err.Number <> 0 Or kids Is Nothing Then Exit Sub
    childCount = kids.Count
    For i = 0 To childCount - 1
        Set child = kids.Item(i)
        WalkTree child, depth + 1, lines, gridCount, visitCount
    Next
End Sub

Sub TryProbePath(sapSession, pathId, ByRef hitCount)
    On Error Resume Next
    Dim ctl, t, rc, cc
    Set ctl = sapSession.FindById(pathId)
    If ctl Is Nothing Then Exit Sub
    hitCount = hitCount + 1
    rc = 0: cc = 0
    rc = CLng(ctl.RowCount)
    cc = CLng(ctl.ColumnCount)
    LogStep "  FOUND: " & ctl.Type & " | " & pathId & " | rows=" & rc & " cols=" & cc
End Sub

Sub ProbeFindByIdDiscovery(sapSession, usrPath, prog, scr, ByRef hitCount)
    Dim cntlNames, ci, shellSuffixes, si, pathId, n, scrNum, prefixes, pi

    LogStep "--- FindById path probing (Children may be empty) ---"

    cntlNames = Array("GRID1", "GRID2", "GRID", "ALV", "ALV_CONTAINER")
    shellSuffixes = Array("", "/shellcont/shell", "/shellcont/shell/collector")
    For ci = 0 To UBound(cntlNames)
        For si = 0 To UBound(shellSuffixes)
            TryProbePath sapSession, usrPath & "/cntl" & cntlNames(ci) & shellSuffixes(si), hitCount
        Next
    Next

    If Len(prog) > 0 And IsNumeric(scr) Then
        scrNum = CLng(scr)
        prefixes = Array("subSUBSCREEN_BODY", "subSUBSCREEN_HEADER")
        For n = scrNum - 5 To scrNum + 5
            For pi = 0 To UBound(prefixes)
                TryProbePath sapSession, usrPath & "/" & prefixes(pi) & ":" & prog & ":" & CStr(n), hitCount
            Next
        Next
    End If

    TryProbePath sapSession, usrPath & "/ctxtS_EBELN-LOW", hitCount
    TryProbePath sapSession, usrPath & "/ctxtS_LIFNR-LOW", hitCount
    TryProbePath sapSession, usrPath & "/lblS_LIFNR-LOW", hitCount
    TryProbePath sapSession, usrPath & "/lblS_EBELN-LOW", hitCount

    LogStep "--- Sample field values ---"
    Dim sampleIds, si2, ctl2, txt2
    sampleIds = Array( _
        usrPath & "/ctxtS_LIFNR-LOW", usrPath & "/ctxtS_EBELN-LOW", _
        usrPath & "/lblS_LIFNR-LOW", usrPath & "/lblS_EBELN-LOW", _
        usrPath & "/cntlGRID1/shellcont/shell/shellcont[1]/shell", _
        usrPath & "/cntlGRID1/shellcont/shell" _
    )
    For si2 = 0 To UBound(sampleIds)
        Set ctl2 = sapSession.FindById(sampleIds(si2))
        If Not ctl2 Is Nothing Then
            txt2 = ""
            On Error Resume Next
            txt2 = Trim(ctl2.Text)
            LogStep "  " & sampleIds(si2) & " => type=" & ctl2.Type & " text=[" & txt2 & "] DefaultTooltip=[" & Trim(ctl2.DefaultTooltip) & "]"
        Else
            LogStep "  MISSING: " & sampleIds(si2)
        End If
    Next

    LogStep "FindById probe hits: " & hitCount
End Sub

Function AttachSapSessionVerbose(application)
    On Error Resume Next
    Dim connection, session, best, connCount, sessCount, activeSess

    connCount = application.Children.Count
    LogStep "connections (application.Children.Count): " & connCount

    Set activeSess = application.ActiveSession
    LogStep "ActiveSession: " & IIf(activeSess Is Nothing, "NO", "YES")
    If Not activeSess Is Nothing Then
        LogStep "  ActiveSession txn: " & activeSess.Info.Transaction
    End If

    If Not IsObject(connection) Then
        If connCount > 0 Then
            Set connection = application.Children(0)
            LogStep "connection via application.Children(0): YES"
        End If
    End If

    If Not IsObject(session) Then
        If IsObject(connection) Then
            sessCount = connection.Children.Count
            LogStep "sessions in connection: " & sessCount
            If sessCount > 0 Then Set session = connection.Children(0)
        End If
    End If

    If IsObject(WScript) Then
        If IsObject(session) Then WScript.ConnectObject session, "on"
        WScript.ConnectObject application, "on"
    End If

    If Not activeSess Is Nothing Then
        Set best = activeSess
    ElseIf IsObject(session) Then
        Set best = session
    End If

    If best Is Nothing Then
        LogStep "RESULT: session NOT attached"
        Set AttachSapSessionVerbose = Nothing
    Else
        LogStep "RESULT: session attached — txn=" & best.Info.Transaction & " program=" & best.Info.Program
        Set AttachSapSessionVerbose = best
    End If
End Function

Sub Main()
    On Error Resume Next
    Dim SapGuiAuto, application, sapSession, usr, lines, gridCount, visitCount, key, usrChildCount, hitCount, prog, scr

    Set lines = CreateObject("Scripting.Dictionary")
    gridCount = 0
    visitCount = 0
    hitCount = 0

    LogStep "=== SAP session attach ==="

    Set SapGuiAuto = GetObject("SAPGUI")
    If SapGuiAuto Is Nothing Then
        LogStep "FAIL: SAPGUI not found"
        WScript.Quit 1
    End If
    Set application = SapGuiAuto.GetScriptingEngine
    If application Is Nothing Then
        LogStep "FAIL: GetScriptingEngine NULL"
        WScript.Quit 2
    End If

    Set sapSession = AttachSapSessionVerbose(application)
    If sapSession Is Nothing Then WScript.Quit 3

    prog = sapSession.Info.Program
    scr = sapSession.Info.ScreenNumber
    LogStep "Window: " & sapSession.FindById("wnd[0]").Text

    Set usr = sapSession.FindById("wnd[0]/usr")
    If usr Is Nothing Then
        LogStep "FAIL: wnd[0]/usr not found"
        WScript.Quit 4
    End If

    usrChildCount = usr.Children.Count
    LogStep "usr.Children.Count: " & usrChildCount & " (0 = use FindById probing only)"

    LogStep "=== Index walk usr.Children(0..n) ==="
    Dim i, child, idxVisits, typeCounts, tKey
    Set typeCounts = CreateObject("Scripting.Dictionary")
    idxVisits = 0
    For i = 0 To usrChildCount - 1
        If i > 400 Then Exit For
        Err.Clear
        Set child = usr.Children(i)
        If Not child Is Nothing Then
            idxVisits = idxVisits + 1
            tKey = child.Type
            If typeCounts.Exists(tKey) Then
                typeCounts(tKey) = typeCounts(tKey) + 1
            Else
                typeCounts.Add tKey, 1
            End If
            If idxVisits <= 20 Then
                WScript.Echo "  [" & i & "] " & child.Type & " | " & child.Id
            End If
        End If
    Next
    If idxVisits > 20 Then LogStep "  ... and " & (idxVisits - 20) & " more children"
    LogStep "Index walk found " & idxVisits & " children"
    LogStep "=== Type histogram ==="
    For Each tKey In typeCounts.Keys
        WScript.Echo "  " & tKey & ": " & typeCounts(tKey)
    Next

    LogStep "=== cntl/grid shell probe from index ==="
    Dim suffixes, si, probeId, ctl, rc, cc
    suffixes = Array("/shellcont/shell", "/shellcont/shell/collector", "/shellcont/shell/shellcont[1]/shell")
    For i = 0 To usrChildCount - 1
        If i > 400 Then Exit For
        Set child = usr.Children(i)
        If Not child Is Nothing Then
            If InStr(LCase(child.Id), "cntl") > 0 Or InStr(LCase(child.Id), "grid") > 0 Then
                LogStep "  cntl child: " & child.Type & " | " & child.Id
                For si = 0 To UBound(suffixes)
                    probeId = child.Id & suffixes(si)
                    Set ctl = sapSession.FindById(probeId)
                    If Not ctl Is Nothing Then
                        rc = 0: cc = 0
                        rc = CLng(ctl.RowCount)
                        cc = CLng(ctl.ColumnCount)
                        LogStep "    -> " & ctl.Type & " rows=" & rc & " cols=" & cc & " | " & probeId
                    End If
                Next
            End If
        End If
    Next

    LogStep "=== Recursive tree walk (top 20 lines) ==="
    WalkTree usr, 0, lines, gridCount, visitCount
    For Each key In lines.Keys
        WScript.Echo key
    Next
    LogStep "Tree walk: visited=" & visitCount & " grid-like=" & gridCount

    ProbeFindByIdDiscovery sapSession, usr.Id, prog, scr, hitCount
End Sub

Call Main()
