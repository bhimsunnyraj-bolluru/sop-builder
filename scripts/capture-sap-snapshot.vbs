' SAP GUI Scripting bridge — VBScript late binding (PowerShell TYPE_E_CANTLOADLIBRARY workaround).
' Usage: cscript //Nologo capture-sap-snapshot.vbs
Option Explicit

Dim gDiscoverMs
gDiscoverMs = 0
Dim gMaxCaptureFields
gMaxCaptureFields = 10
Dim gDbgVisited
gDbgVisited = 0
Dim gDbgGridCandidates
gDbgGridCandidates = 0
Dim gDbgGridCells
gDbgGridCells = 0
Dim gDbgTypeSamples
Set gDbgTypeSamples = Nothing
Dim gSessionDbg
Set gSessionDbg = Nothing

Function JsonEscape(s)
    If IsNull(s) Or IsEmpty(s) Then JsonEscape = "": Exit Function
    s = CStr(s)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function

Function SafeStr(v)
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Then
        SafeStr = ""
    Else
        SafeStr = CStr(v)
    End If
End Function

Function JsonNum(v)
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Then
        JsonNum = "0"
    ElseIf IsNumeric(v) Then
        JsonNum = CStr(CLng(v))
    Else
        JsonNum = "0"
    End If
End Function

Sub SessionDbgInit()
    Set gSessionDbg = CreateObject("Scripting.Dictionary")
End Sub

Sub SessionDbgSet(key, value)
    On Error Resume Next
    If gSessionDbg Is Nothing Then SessionDbgInit()
    If gSessionDbg.Exists(key) Then
        gSessionDbg(key) = value
    Else
        gSessionDbg.Add key, value
    End If
End Sub

Function SessionDbgJson()
    Dim json, k, v, first, vs
    json = ""
    first = True
    If gSessionDbg Is Nothing Then
        SessionDbgJson = ""
        Exit Function
    End If
    For Each k In gSessionDbg.Keys
        v = gSessionDbg(k)
        If Not first Then json = json & ","
        first = False
        vs = LCase(CStr(v))
        If vs = "true" Or vs = "false" Then
            json = json & """" & JsonEscape(CStr(k)) & """:" & vs
        ElseIf IsNumeric(v) Then
            json = json & """" & JsonEscape(CStr(k)) & """:" & CStr(v)
        Else
            json = json & """" & JsonEscape(CStr(k)) & """:""" & JsonEscape(CStr(v)) & """"
        End If
    Next
    SessionDbgJson = json
End Function

Function HasVisibleBounds(ctl)
    On Error Resume Next
    HasVisibleBounds = (ReadIntCoord(ctl, "ScreenWidth", "Width") > 0 And ReadIntCoord(ctl, "ScreenHeight", "Height") > 0)
End Function

Function PathTail(pathId)
    If Len(pathId) = 0 Then
        PathTail = ""
    Else
        PathTail = Mid(pathId, InStrRev(pathId, "/") + 1)
    End If
End Function

Function VerifyControlForPath(ctl, pathId)
    On Error Resume Next
    Dim cid, tail, reqTail, lt, lPath
    VerifyControlForPath = False
    If ctl Is Nothing Or Len(pathId) = 0 Then Exit Function
    cid = Trim(ctl.Id)
    If Len(cid) = 0 Then Exit Function
    reqTail = PathTail(pathId)
    tail = PathTail(cid)
    If StrComp(tail, reqTail, vbTextCompare) <> 0 Then Exit Function
    lt = LCase(ctl.Type)
    lPath = LCase(pathId)
    If InStr(lPath, "/lbl") > 0 Then
        If InStr(lt, "label") = 0 Then Exit Function
    ElseIf InStr(lPath, "/ctxt") > 0 Or InStr(lPath, "/txt") > 0 Or InStr(lPath, "/cmb") > 0 Then
        If InStr(lt, "text") = 0 And InStr(lt, "combo") = 0 And InStr(lt, "check") = 0 Then Exit Function
    ElseIf InStr(lPath, "shell") > 0 Or InStr(lPath, "/cntl") > 0 Or InStr(lPath, "grid") > 0 Then
        If InStr(lt, "textfield") > 0 Then Exit Function
    End If
    If InStr(lPath, "shell") > 0 Or InStr(lPath, "/cntlgrid") > 0 Or InStr(lPath, "/cntal") > 0 Then
        If Not HasVisibleBounds(ctl) And InStr(lt, "grid") = 0 And InStr(lt, "table") = 0 And InStr(lt, "shell") = 0 Then Exit Function
    End If
    VerifyControlForPath = True
End Function

Function FindByIdVerified(session, pathId)
    On Error Resume Next
    Dim ctl
    Set FindByIdVerified = Nothing
    If Len(pathId) = 0 Then Exit Function
    Err.Clear
    Set ctl = Nothing
    Set ctl = session.FindById(pathId)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    If VerifyControlForPath(ctl, pathId) Then Set FindByIdVerified = ctl
End Function

Function CleanSapFieldLabel(raw)
    Dim s, openPos, closePos
    s = Trim(raw)
    If Len(s) = 0 Then
        CleanSapFieldLabel = ""
        Exit Function
    End If
    openPos = InStrRev(s, "(")
    closePos = InStrRev(s, ")")
    If openPos > 0 And closePos > openPos Then
        s = Trim(Left(s, openPos - 1))
    End If
    CleanSapFieldLabel = Trim(Replace(s, ":", ""))
End Function

Function ReadFieldLabel(ctl, ctlType)
    On Error Resume Next
    Dim lt
    lt = LCase(Trim(ctlType))
    If InStr(lt, "ctextfield") > 0 Then
        ReadFieldLabel = CleanSapFieldLabel(ctl.DefaultTooltip)
        Exit Function
    End If
    ReadFieldLabel = CleanSapFieldLabel(ctl.Tooltip)
End Function

Function ReadFieldValue(fld)
    On Error Resume Next
    Dim v
    v = fld.Text
    If Len(Trim(v)) = 0 Then v = fld.DisplayedText
    If Len(Trim(v)) = 0 Then v = fld.Key
    If Len(Trim(v)) = 0 Then v = fld.Value
    ReadFieldValue = Trim(v)
End Function

Function CollectionIndexBase(kids)
    On Error Resume Next
    Dim child
    CollectionIndexBase = 0
    If kids Is Nothing Then Exit Function
    Err.Clear
    Set child = Nothing
    Set child = kids.Item(CLng(0))
    If Err.Number = 0 And Not child Is Nothing Then
        CollectionIndexBase = 0
        Exit Function
    End If
    Err.Clear
    Set child = Nothing
    Set child = kids.Item(CLng(1))
    If Err.Number = 0 And Not child Is Nothing Then
        CollectionIndexBase = 1
        Exit Function
    End If
    Err.Clear
    CollectionIndexBase = 0
End Function

Function ChildAt(parent, kids, index)
    On Error Resume Next
    Dim child
    Set ChildAt = Nothing
    If kids Is Nothing Then Exit Function
    Err.Clear
    Set child = Nothing
    Set child = kids.Item(CLng(index))
    If Err.Number = 0 And Not child Is Nothing Then
        Set ChildAt = child
        Exit Function
    End If
    Err.Clear
    Set child = Nothing
    Set child = kids(CLng(index))
    If Err.Number = 0 And Not child Is Nothing Then
        Set ChildAt = child
        Exit Function
    End If
    Err.Clear
    If Not parent Is Nothing Then
        Set child = Nothing
        Set child = parent.Children(CLng(index))
        If Err.Number = 0 And Not child Is Nothing Then Set ChildAt = child
    End If
    Err.Clear
End Function

Function ResolveControlName(fld)
    On Error Resume Next
    Dim n, id, tail
    n = Trim(fld.Name)
    id = Trim(fld.Id)
    If Len(n) > 0 Then
        ResolveControlName = n
        Exit Function
    End If
    If Len(id) = 0 Then
        ResolveControlName = "field"
        Exit Function
    End If
    tail = Mid(id, InStrRev(id, "/") + 1)
    tail = Replace(tail, "ctxt", "", 1, 1, vbTextCompare)
    tail = Replace(tail, "txt", "", 1, 1, vbTextCompare)
    If Len(tail) > 0 Then
        ResolveControlName = tail
    Else
        ResolveControlName = id
    End If
End Function

Function IsTechnicalType(t)
    Dim lt
    lt = LCase(t)
    Select Case lt
        Case "guimenubar", "guitoolbar", "guititlebar", "guistatuspane", "guiokcodefield", _
             "guinavigationpanel", "guisplitter", "guiscrollcontainer", "guicontainershell", "guishell"
            IsTechnicalType = True
        Case Else
            IsTechnicalType = False
    End Select
End Function

Function ReadIntCoord(fld, primaryProp, fallbackProp)
    On Error Resume Next
    Dim v
    v = 0
    If LCase(primaryProp) = "screenleft" Then v = fld.ScreenLeft
    If LCase(primaryProp) = "screentop" Then v = fld.ScreenTop
    If LCase(primaryProp) = "screenwidth" Then v = fld.ScreenWidth
    If LCase(primaryProp) = "screenheight" Then v = fld.ScreenHeight
    If (Not IsNumeric(v) Or CLng(v) = 0) And Len(fallbackProp) > 0 Then
        If LCase(fallbackProp) = "left" Then v = fld.Left
        If LCase(fallbackProp) = "top" Then v = fld.Top
        If LCase(fallbackProp) = "width" Then v = fld.Width
        If LCase(fallbackProp) = "height" Then v = fld.Height
    End If
    If Not IsNumeric(v) Then
        ReadIntCoord = 0
    Else
        ReadIntCoord = CLng(v)
    End If
End Function

Function IsEditableFieldId(id)
    If InStr(1, id, "ctxt", vbTextCompare) > 0 Then
        IsEditableFieldId = True
        Exit Function
    End If
    If InStr(1, id, "/txt", vbTextCompare) > 0 Then
        IsEditableFieldId = True
        Exit Function
    End If
    If InStr(1, id, "cmb", vbTextCompare) > 0 Then
        IsEditableFieldId = True
        Exit Function
    End If
    If InStr(1, id, "pwd", vbTextCompare) > 0 Then
        IsEditableFieldId = True
        Exit Function
    End If
    If InStr(1, id, "chk", vbTextCompare) > 0 Then
        IsEditableFieldId = True
        Exit Function
    End If
    If InStr(1, id, "/cell[", vbTextCompare) > 0 Then
        IsEditableFieldId = True
        Exit Function
    End If
    IsEditableFieldId = False
End Function

Function IsCapturedFieldId(id)
    IsCapturedFieldId = IsEditableFieldId(id)
End Function

Function IsEditableFieldType(t)
    Dim lt
    lt = LCase(t)
    If InStr(lt, "textfield") > 0 Then
        IsEditableFieldType = True
        Exit Function
    End If
    If InStr(lt, "textedit") > 0 Then
        IsEditableFieldType = True
        Exit Function
    End If
    If InStr(lt, "combobox") > 0 Then
        IsEditableFieldType = True
        Exit Function
    End If
    If InStr(lt, "checkbox") > 0 Then
        IsEditableFieldType = True
        Exit Function
    End If
    If InStr(lt, "radiobutton") > 0 Then
        IsEditableFieldType = True
        Exit Function
    End If
    IsEditableFieldType = False
End Function

Function IsGenericColumnTitle(title)
    Dim t
    t = UCase(Trim(title))
    If Len(t) = 0 Then
        IsGenericColumnTitle = True
        Exit Function
    End If
    If t = "SDITEM" Or t = "SD" Or t = "ITEM" Or t = "FIELD" Then
        IsGenericColumnTitle = True
        Exit Function
    End If
    If Left(t, 6) = "SDITEM" Then
        IsGenericColumnTitle = True
        Exit Function
    End If
    IsGenericColumnTitle = False
End Function

Function CaptureLimitReached(count)
    CaptureLimitReached = (gMaxCaptureFields > 0 And count >= gMaxCaptureFields)
End Function

Sub AppendControlNode(ctl, t, id, txt, nm, ByRef arr, ByRef count, labelOverride)
    Dim node, fieldLabel, rawTip
    If CaptureLimitReached(count) Then Exit Sub
    If IsEmpty(labelOverride) Then labelOverride = ""
    fieldLabel = ReadFieldLabel(ctl, t)
    If Len(fieldLabel) = 0 Then fieldLabel = Trim(labelOverride)
    If IsGenericColumnTitle(fieldLabel) Then fieldLabel = ""
    ReDim Preserve arr(count)
    Set node = CreateObject("Scripting.Dictionary")
    node.Add "id", id
    node.Add "type", t
    node.Add "name", nm
    node.Add "text", txt
    node.Add "value", txt
    node.Add "tooltip", ctl.Tooltip
    node.Add "left", ReadIntCoord(ctl, "ScreenLeft", "Left")
    node.Add "top", ReadIntCoord(ctl, "ScreenTop", "Top")
    node.Add "width", ReadIntCoord(ctl, "ScreenWidth", "Width")
    node.Add "height", ReadIntCoord(ctl, "ScreenHeight", "Height")
    If Len(fieldLabel) > 0 Then node.Add "label", fieldLabel
    If InStr(LCase(t), "ctextfield") > 0 Then
        rawTip = Trim(ctl.DefaultTooltip)
        If Len(rawTip) > 0 Then node.Add "defaultTooltip", rawTip
    End If
    Set arr(count) = node
    count = count + 1
End Sub

Sub AppendSyntheticFieldNode(fieldId, fieldType, fieldValue, fieldName, fieldLabel, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim node, lbl
    If CaptureLimitReached(count) Then Exit Sub
    If Len(fieldId) = 0 Or Len(fieldValue) = 0 Then Exit Sub
    If seen.Exists(fieldId) Then Exit Sub
    seen.Add fieldId, True
    lbl = Trim(fieldLabel)
    If IsGenericColumnTitle(lbl) Then lbl = ""
    ReDim Preserve arr(count)
    Set node = CreateObject("Scripting.Dictionary")
    node.Add "id", fieldId
    node.Add "type", fieldType
    node.Add "name", fieldName
    node.Add "text", fieldValue
    node.Add "value", fieldValue
    node.Add "tooltip", ""
    node.Add "left", 0
    node.Add "top", 0
    node.Add "width", 0
    node.Add "height", 0
    If Len(lbl) > 0 Then node.Add "label", lbl
    Set arr(count) = node
    count = count + 1
End Sub

Function GetGridColumnTitle(grid, col)
    On Error Resume Next
    Dim title, colOrd, item
    title = ""
    Err.Clear
    title = Trim(grid.GetDisplayedColumnTitle(col))
    If Len(title) > 0 And Not IsGenericColumnTitle(title) Then
        GetGridColumnTitle = title
        Exit Function
    End If
    Err.Clear
    title = Trim(grid.GetColumnTitle(col))
    If Len(title) > 0 And Not IsGenericColumnTitle(title) Then
        GetGridColumnTitle = title
        Exit Function
    End If
    Err.Clear
    Set colOrd = grid.ColumnOrder
    If Not colOrd Is Nothing Then
        If colOrd.Count > col Then
            Set item = colOrd.Item(col)
            If Not item Is Nothing Then
                title = Trim(item.Title)
                If Len(title) = 0 Then title = Trim(item.Name)
                If Len(title) > 0 And Not IsGenericColumnTitle(title) Then
                    GetGridColumnTitle = title
                    Exit Function
                End If
            End If
        End If
    End If
    GetGridColumnTitle = "Column " & CStr(col + 1)
End Function

Sub DbgNoteControlType(ctl)
    On Error Resume Next
    Dim t
    If ctl Is Nothing Then Exit Sub
    gDbgVisited = gDbgVisited + 1
    If gDbgTypeSamples Is Nothing Then Set gDbgTypeSamples = CreateObject("Scripting.Dictionary")
    t = Trim(ctl.Type)
    If Len(t) > 0 Then
        If Not gDbgTypeSamples.Exists(t) Then
            If gDbgTypeSamples.Count < 20 Then gDbgTypeSamples.Add t, True
        End If
    End If
End Sub

Function GridLooksReadable(grid)
    On Error Resume Next
    Dim rc, cc, lt
    GridLooksReadable = False
    If grid Is Nothing Then Exit Function
    lt = LCase(grid.Type)
    If InStr(lt, "grid") > 0 Or InStr(lt, "table") > 0 Then
        GridLooksReadable = True
        Exit Function
    End If
    rc = CLng(grid.RowCount)
    cc = CLng(grid.ColumnCount)
    If rc <= 0 Or cc <= 0 Then Exit Function
    Err.Clear
    grid.GetCellValue 0, 0
    GridLooksReadable = (Err.Number = 0)
End Function

Sub ResolveActualGridControl(root, ByRef gridOut, depth)
    On Error Resume Next
    Dim rowCount, colCount, kids, i, child, gv
    If root Is Nothing Or depth > 18 Then Exit Sub
    If Not gridOut Is Nothing Then Exit Sub

    If GridLooksReadable(root) Then
        gDbgGridCandidates = gDbgGridCandidates + 1
        Set gridOut = root
        Exit Sub
    End If

    Err.Clear
    Set gv = root.GetGridView()
    If Not gv Is Nothing Then
        If GridLooksReadable(gv) Then
            gDbgGridCandidates = gDbgGridCandidates + 1
            Set gridOut = gv
            Exit Sub
        End If
    End If

    Set kids = root.Children
    If Err.Number <> 0 Or kids Is Nothing Then Exit Sub
    For i = 0 To kids.Count - 1
        Err.Clear
        Set child = kids.Item(i)
        If Not child Is Nothing Then ResolveActualGridControl child, gridOut, depth + 1
        If Not gridOut Is Nothing Then Exit Sub
    Next
End Sub

Function ReadGridCellValue(grid, r, c)
    On Error Resume Next
    Dim val, cellObj, colOrd, item, colName
    val = ""
    val = Trim(grid.GetCellValue(r, c))
    If Len(val) > 0 Then ReadGridCellValue = val: Exit Function
    val = Trim(grid.GetCellText(r, c))
    If Len(val) > 0 Then ReadGridCellValue = val: Exit Function
    Err.Clear
    val = Trim(grid.GetDisplayedCellValue(r, c))
    If Len(val) > 0 Then ReadGridCellValue = val: Exit Function
    Set colOrd = grid.ColumnOrder
    If Not colOrd Is Nothing Then
        If colOrd.Count > c Then
            Set item = colOrd.Item(c)
            If Not item Is Nothing Then
                colName = Trim(item.Name)
                If Len(colName) > 0 Then
                    val = Trim(grid.GetCellValue(r, colName))
                    If Len(val) > 0 Then ReadGridCellValue = val: Exit Function
                    val = Trim(grid.GetCellText(r, colName))
                    If Len(val) > 0 Then ReadGridCellValue = val: Exit Function
                End If
            End If
        End If
    End If
    Set cellObj = Nothing
    Set cellObj = grid.GetCell(r, c)
    If Not cellObj Is Nothing Then val = ReadFieldValue(cellObj)
    ReadGridCellValue = val
End Function

Sub CaptureGridValues(session, grid, gridId, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim rowCount, colCount, r, c, maxR, maxC, val, colTitle, fieldId, nm, firstRow, lastRow
    If grid Is Nothing Or Len(gridId) = 0 Then Exit Sub

    rowCount = CLng(grid.RowCount)
    If rowCount <= 0 Then rowCount = CLng(grid.VisibleRowCount)
    colCount = CLng(grid.ColumnCount)
    If rowCount <= 0 Or colCount <= 0 Then Exit Sub

    firstRow = 0
    lastRow = rowCount - 1
    Err.Clear
    firstRow = CLng(grid.FirstVisibleRow)
    lastRow = CLng(grid.LastVisibleRow)
    If Err.Number <> 0 Or lastRow < firstRow Then
        firstRow = 0
        lastRow = rowCount - 1
    End If
    maxR = lastRow
    If maxR > firstRow + 19 Then maxR = firstRow + 19
    maxC = colCount - 1
    If maxC > 30 Then maxC = 30

    For r = firstRow To maxR
        For c = 0 To maxC
            val = ReadGridCellValue(grid, r, c)
            If Len(val) > 0 Then
                colTitle = GetGridColumnTitle(grid, c)
                fieldId = gridId & "/cell[" & r & "," & c & "]"
                nm = colTitle
                If Len(nm) = 0 Then nm = "R" & r & "C" & c
                AppendSyntheticFieldNode fieldId, "GuiGridViewCell", val, nm, colTitle, arr, count, seen
                gDbgGridCells = gDbgGridCells + 1
            End If
        Next
    Next
End Sub

Sub TryCaptureGridControl(session, ctl, gridId, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim grid
    Set grid = Nothing
    ResolveActualGridControl ctl, grid, 0
    If grid Is Nothing Then Exit Sub
    If Len(grid.Id) > 0 Then
        CaptureGridValues session, grid, grid.Id, arr, count, seen
    Else
        CaptureGridValues session, grid, gridId, arr, count, seen
    End If
End Sub

Sub ProbeAlvShellPaths(session, usrPath, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim usr, i, child, id, suffixes, si, probeId, ctl, j, subChild
    suffixes = Array("/shellcont/shell", "/shellcont/shell/collector", "/shellcont/shell/shellcont[0]/shell", "/shellcont/shell/shellcont/shell")
    Set usr = session.FindById(usrPath)
    If usr Is Nothing Then Exit Sub

    For i = 0 To 120
        Err.Clear
        Set child = Nothing
        Set child = usr.Children(i)
        If Err.Number <> 0 Then Exit For
        If child Is Nothing Then Exit For
        id = LCase(child.Id)
        If InStr(id, "cntl") > 0 Or InStr(id, "grid") > 0 Or InStr(id, "alv") > 0 Then
            For si = 0 To UBound(suffixes)
                probeId = child.Id & suffixes(si)
                Set ctl = FindByIdVerified(session, probeId)
                If Not ctl Is Nothing Then
                    DbgNoteControlType ctl
                    TryCaptureGridControl session, ctl, probeId, arr, count, seen
                End If
            Next
            TryCaptureGridControl session, child, child.Id, arr, count, seen
        End If
    Next
End Sub

Sub ProbePathIfExists(session, probePath, ByRef arr, ByRef count, ByRef seen, ByRef foundCount)
    On Error Resume Next
    Dim ctl, lt
    If Len(probePath) = 0 Then Exit Sub
    Set ctl = FindByIdVerified(session, probePath)
    If ctl Is Nothing Then Exit Sub
    foundCount = foundCount + 1
    DbgNoteControlType ctl
    ProcessSapControl session, ctl, arr, count, seen
    TryCaptureGridControl session, ctl, probePath, arr, count, seen
    lt = LCase(ctl.Type)
    If InStr(lt, "table") > 0 Then
        CaptureTableFields session, probePath, ctl, arr, count, seen
    End If
End Sub

Sub ProbeCntlGridPathsMinimal(session, usrPath, ByRef arr, ByRef count, ByRef seen, ByRef foundCount)
    On Error Resume Next
    Dim paths, i
    paths = Array( _
        usrPath & "/cntlGRID1/shellcont/shell/shellcont[1]/shell", _
        usrPath & "/cntlGRID1/shellcont/shell/collector", _
        usrPath & "/cntlGRID1/shellcont/shell", _
        usrPath & "/cntlGRID2/shellcont/shell/shellcont[1]/shell", _
        usrPath & "/cntlALV/shellcont/shell/collector", _
        usrPath & "/cntlCUSTOM_CONTROL/shellcont/shell" _
    )
    For i = 0 To UBound(paths)
        ProbePathIfExists session, paths(i), arr, count, seen, foundCount
    Next
End Sub

Sub ProbeCntlGridPathsFast(session, usrPath, ByRef arr, ByRef count, ByRef seen, ByRef foundCount)
    On Error Resume Next
    Dim cntlNames, ci, shellSuffixes, si, probePath
    cntlNames = Array( _
        "GRID1", "GRID2", "GRID", "ALV", "ALV_CONTAINER", "ALV_TREE", _
        "CONTAINER", "CUSTOM_CONTROL", "LIST", "TABLE" _
    )
    shellSuffixes = Array( _
        "", "/shellcont/shell", "/shellcont/shell/collector", _
        "/shellcont/shell/shellcont[1]/shell", "/shellcont/shell/shellcont[0]/shell", _
        "/shellcont/shell/shellcont/shell" _
    )

    For ci = 0 To UBound(cntlNames)
        For si = 0 To UBound(shellSuffixes)
            probePath = usrPath & "/cntl" & cntlNames(ci) & shellSuffixes(si)
            ProbePathIfExists session, probePath, arr, count, seen, foundCount
        Next
    Next
End Sub

Sub ProbeCntlPathsByFindById(session, usrPath, ByRef arr, ByRef count, ByRef seen, ByRef foundCount)
    On Error Resume Next
    Dim ci, probePath
    ProbeCntlGridPathsFast session, usrPath, arr, count, seen, foundCount
    For ci = 0 To 12
        probePath = usrPath & "/cntl" & CStr(ci)
        ProbePathIfExists session, probePath, arr, count, seen, foundCount
        probePath = usrPath & "/cntlGRID" & CStr(ci)
        ProbePathIfExists session, probePath, arr, count, seen, foundCount
    Next
End Sub

Sub ProbeSubscreensByFindById(session, usrPath, prog, scr, ByRef arr, ByRef count, ByRef seen, ByRef foundCount, span)
    On Error Resume Next
    Dim prefixes, pi, n, scrNum, probePath, scrTry
    If Len(prog) = 0 Then Exit Sub
    If span <= 0 Then span = 3

    prefixes = Array("subSUBSCREEN_BODY", "subSUBSCREEN_HEADER", "subSCREEN", "subSUBSCREEN")
    If IsNumeric(scr) Then scrNum = CLng(scr)
    For n = scrNum - span To scrNum + span
        scrTry = CStr(n)
        For pi = 0 To UBound(prefixes)
            probePath = usrPath & "/" & prefixes(pi) & ":" & prog & ":" & scrTry
            ProbePathIfExists session, probePath, arr, count, seen, foundCount
        Next
    Next
End Sub

Sub ProbeSelectionFieldsByFindById(session, usrPath, ByRef arr, ByRef count, ByRef seen, ByRef foundCount)
    On Error Resume Next
    Dim suffixes, i, prefixes, pi, fieldId, fld
    prefixes = Array("ctxt")
    suffixes = Array( _
        "S_EBELN-LOW", "S_EBELN-HIGH", "S_LIFNR-LOW", "S_LIFNR-HIGH", _
        "S_MATNR-LOW", "S_MATNR-HIGH", "S_WERKS-LOW", "S_BUKRS-LOW", _
        "S_BEDAT-LOW", "S_BEDAT-HIGH", "S_BSART-LOW", "S_EKORG-LOW", _
        "S_EKORG-HIGH", "S_EKGRP-LOW", "S_EKGRP-HIGH", "S_STATU-LOW", _
        "S_AEDAT-LOW", "S_AEDAT-HIGH" _
    )
    For i = 0 To UBound(suffixes)
        For pi = 0 To UBound(prefixes)
            fieldId = usrPath & "/" & prefixes(pi) & suffixes(i)
            If Not seen.Exists(fieldId) Then
                Set fld = FindByIdVerified(session, fieldId)
                If Not fld Is Nothing Then
                    foundCount = foundCount + 1
                    ProcessSapControl session, fld, arr, count, seen
                End If
            End If
        Next
    Next
End Sub

Sub ProbeSelectionFieldsInContainer(session, containerPath, ByRef arr, ByRef count, ByRef seen, ByRef foundCount)
    On Error Resume Next
    Dim suffixes, i, prefixes, pi, fieldId, fld
    If Len(containerPath) = 0 Then Exit Sub
    prefixes = Array("ctxt")
    suffixes = Array( _
        "S_EBELN-LOW", "S_EBELN-HIGH", "S_LIFNR-LOW", "S_LIFNR-HIGH", _
        "S_MATNR-LOW", "S_MATNR-HIGH", "S_WERKS-LOW", "S_BUKRS-LOW", _
        "S_BEDAT-LOW", "S_BEDAT-HIGH", "S_BSART-LOW", "S_EKORG-LOW", _
        "S_EKORG-HIGH", "S_EKGRP-LOW", "S_EKGRP-HIGH", "S_STATU-LOW", _
        "S_AEDAT-LOW", "S_AEDAT-HIGH" _
    )
    For i = 0 To UBound(suffixes)
        For pi = 0 To UBound(prefixes)
            fieldId = containerPath & "/" & prefixes(pi) & suffixes(i)
            If Not seen.Exists(fieldId) Then
                Set fld = FindByIdVerified(session, fieldId)
                If Not fld Is Nothing Then
                    foundCount = foundCount + 1
                    ProcessSapControl session, fld, arr, count, seen
                End If
            End If
        Next
    Next
End Sub

Sub ProbeFindAllOnContainer(session, container, ByRef arr, ByRef count, ByRef seen, ByRef foundCount)
    On Error Resume Next
    Dim names, types, ni, ti, coll, i, ctl
    If container Is Nothing Then Exit Sub
    names = Array( _
        "S_EBELN", "S_LIFNR", "S_MATNR", "S_WERKS", "S_BUKRS", "S_BEDAT", _
        "S_EKORG", "S_EKGRP", "S_BSART", _
        "EBELN", "LIFNR", "MATNR", "WERKS", "BUKRS", "EKORG", "EKGRP", "BSART", _
        "AUART", "KUNNR", "KUNWE" _
    )
    types = Array("GuiTextField", "GuiCTextField", "GuiComboBox", "GuiCheckBox")
    For ni = 0 To UBound(names)
        For ti = 0 To UBound(types)
            Err.Clear
            Set coll = container.FindAllByName(names(ni), types(ti))
            If Err.Number = 0 And Not coll Is Nothing Then
                For i = 0 To coll.Count - 1
                    Set ctl = coll.Item(i)
                    If Not ctl Is Nothing Then
                        foundCount = foundCount + 1
                        ProcessSapControl session, ctl, arr, count, seen
                    End If
                Next
            End If
        Next
    Next
End Sub

Sub DiscoverByPathProbing(session, usrPath, prog, scr, ByRef arr, ByRef count, ByRef seen, lightMode)
    On Error Resume Next
    Dim usr, foundCount, usrChildCount, subSpan
    foundCount = 0
    usrChildCount = -1
    If IsEmpty(lightMode) Then lightMode = False

    Set usr = session.FindById(usrPath)
    If Not usr Is Nothing Then
        Err.Clear
        usrChildCount = usr.Children.Count
        SessionDbgSet "usrChildrenCount", usrChildCount
        If Not lightMode Then ProbeFindAllOnContainer session, usr, arr, count, seen, foundCount
    End If

    If lightMode Then
        ProbeSelectionFieldsByFindById session, usrPath, arr, count, seen, foundCount
    Else
        ProbeCntlPathsByFindById session, usrPath, arr, count, seen, foundCount
        subSpan = 3
        If StrComp(prog, "SAPMV45A", vbTextCompare) = 0 Then subSpan = 15
        ProbeSubscreensByFindById session, usrPath, prog, scr, arr, count, seen, foundCount, subSpan
        ProbeSelectionFieldsByFindById session, usrPath, arr, count, seen, foundCount
        ProbeAlvShellPathsByFindById session, usrPath, arr, count, seen, foundCount
    End If

    SessionDbgSet "pathProbeHits", foundCount
End Sub

Sub ProbeShellPathsForCntl(session, cntlId, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim suffixes, si, probeId, ctl
    If Len(cntlId) = 0 Then Exit Sub
    suffixes = Array("/shellcont/shell", "/shellcont/shell/collector", "/shellcont/shell/shellcont[1]/shell")
    For si = 0 To UBound(suffixes)
        probeId = cntlId & suffixes(si)
        Set ctl = FindByIdVerified(session, probeId)
        If Not ctl Is Nothing Then
            ProcessSapControl session, ctl, arr, count, seen
            TryCaptureGridControl session, ctl, probeId, arr, count, seen
        End If
    Next
End Sub

Sub DiscoverUsrChildrenByIndex(session, usrPath, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim usr, i, child, maxI, childCount, cid, lt, probeChild
    Set usr = session.FindById(usrPath)
    If usr Is Nothing Then Exit Sub
    childCount = usr.Children.Count
    SessionDbgSet "usrChildrenCount", childCount
    If childCount <= 0 Then Exit Sub
    Set probeChild = usr.Children(0)
    If probeChild Is Nothing Or Len(Trim(probeChild.Id)) = 0 Then
        SessionDbgSet "usrChildrenIndexed", False
        Exit Sub
    End If
    SessionDbgSet "usrChildrenIndexed", True
    maxI = childCount - 1
    If maxI > 400 Then maxI = 400
    For i = 0 To maxI
        Err.Clear
        Set child = Nothing
        Set child = usr.Children(i)
        If Err.Number <> 0 Then
            Err.Clear
            Set child = usr.Children.Item(i)
        End If
        If Not child Is Nothing Then
            ProcessSapControl session, child, arr, count, seen
            cid = LCase(child.Id)
            lt = LCase(child.Type)
            If InStr(cid, "cntl") > 0 Or InStr(cid, "grid") > 0 Or InStr(cid, "alv") > 0 Then
                ProbeShellPathsForCntl session, child.Id, arr, count, seen
            End If
            If InStr(lt, "table") > 0 Or InStr(cid, "/tbl") > 0 Then
                CaptureTableFields session, child.Id, child, arr, count, seen
            End If
            TryCaptureGridControl session, child, child.Id, arr, count, seen
        End If
    Next
End Sub

Sub ProbeAlvShellPathsByFindById(session, usrPath, ByRef arr, ByRef count, ByRef seen, ByRef foundCount)
    ProbeCntlGridPathsFast session, usrPath, arr, count, seen, foundCount
End Sub

Sub WalkAllChildren(session, ctl, ByRef arr, ByRef count, ByRef seen, depth)
    On Error Resume Next
    Dim kids, i, child, childCount, childId
    If ctl Is Nothing Or depth > 48 Then Exit Sub

    DbgNoteControlType ctl
    ProcessSapControl session, ctl, arr, count, seen
    If Len(ctl.Id) > 0 Then TryCaptureGridControl session, ctl, ctl.Id, arr, count, seen

    Set kids = ctl.Children
    If Err.Number <> 0 Or kids Is Nothing Then Exit Sub
    childCount = kids.Count
    For i = 0 To childCount - 1
        Err.Clear
        Set child = Nothing
        Set child = kids(i)
        If Err.Number <> 0 Then
            Err.Clear
            Set child = kids.Item(i)
        End If
        If Not child Is Nothing Then
            childId = child.Id
            If Len(childId) > 0 Then
                WalkAllChildren session, child, arr, count, seen, depth + 1
            Else
                WalkAllChildren session, child, arr, count, seen, depth + 1
            End If
        End If
    Next
End Sub

Sub DiscoverGenericFallback(session, usrPath, prog, scr, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim visited, usrCtl
    Set usrCtl = session.FindById(usrPath)
    If Not usrCtl Is Nothing Then ProbeFindByNameFields session, usrCtl, arr, count, seen
    ProbeAlvShellPaths session, usrPath, arr, count, seen
    If Len(prog) > 0 And Len(scr) > 0 Then
        Set visited = CreateObject("Scripting.Dictionary")
        ProbeSubscreenContainers session, usrPath, prog, scr, visited, arr, count, seen
    End If
End Sub

Sub ProcessSapControl(session, ctl, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim t, lt, txt, nm, id
    If ctl Is Nothing Then Exit Sub
    If CaptureLimitReached(count) Then Exit Sub

    DbgNoteControlType ctl
    t = ctl.Type
    lt = LCase(t)
    id = ctl.Id
    If Len(id) = 0 Then Exit Sub

    If IsTechnicalType(t) Or seen.Exists(id) Then Exit Sub
    If StrComp(t, "GuiCTextField", vbTextCompare) <> 0 Then Exit Sub

    txt = ReadFieldValue(ctl)
    If Len(Trim(txt)) = 0 Then Exit Sub

    nm = ResolveControlName(ctl)
    seen.Add id, True
    AppendControlNode ctl, t, id, txt, nm, arr, count, ""
End Sub

Sub AppendFieldById(session, fieldId, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim fld, txt, nm, t
    If Len(fieldId) = 0 Then Exit Sub
    If seen.Exists(fieldId) Then Exit Sub
    Set fld = FindByIdVerified(session, fieldId)
    If fld Is Nothing Then Exit Sub
    txt = ReadFieldValue(fld)
    t = fld.Type
    nm = ResolveControlName(fld)
    If IsTechnicalType(t) Then Exit Sub
    If Len(txt) = 0 And Not HasVisibleBounds(fld) And Len(ReadFieldLabel(fld, t)) = 0 Then Exit Sub
    seen.Add fieldId, True
    AppendControlNode fld, t, fieldId, txt, nm, arr, count, ""
End Sub

Sub DiscoverByIdBfs(session, startId, ByRef visited, ByRef arr, ByRef count, ByRef seen, depth)
    On Error Resume Next
    Dim ctl, kids, i, child, childId, childCount
    If depth > 32 Then Exit Sub
    If Len(startId) = 0 Then Exit Sub
    If visited.Exists(startId) Then Exit Sub
    visited.Add startId, True

    Set ctl = session.FindById(startId)
    If ctl Is Nothing Then Exit Sub

    ProcessSapControl session, ctl, arr, count, seen

    Set kids = ctl.Children
    If Err.Number <> 0 Then Exit Sub
    childCount = kids.Count
    If childCount <= 0 Then Exit Sub

    For i = 0 To childCount - 1
        Err.Clear
        Set child = Nothing
        Set child = kids(i)
        If Err.Number <> 0 Then
            Err.Clear
            Set child = kids.Item(i)
        End If
        If child Is Nothing Then
            ' skip
        Else
            childId = child.Id
            If Len(childId) > 0 Then
                DiscoverByIdBfs session, childId, visited, arr, count, seen, depth + 1
            Else
                ProcessSapControl session, child, arr, count, seen
            End If
        End If
    Next
End Sub

Sub TryAppendFieldPath(session, fieldId, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim fld
    If Len(fieldId) = 0 Then Exit Sub
    If seen.Exists(fieldId) Then Exit Sub
    Set fld = session.FindById(fieldId)
    If fld Is Nothing Then Exit Sub
    AppendFieldById session, fieldId, arr, count, seen
End Sub

Function GetStandardFieldSuffixes()
    GetStandardFieldSuffixes = Array( _
        "VBAK-AUART", "VBAK-VKORG", "VBAK-VTWEG", "VBAK-SPART", _
        "VBCOM-VKORG", "VBCOM-VTWEG", "VBCOM-SPART", _
        "VBAK-KUNNR", "VBAK-KUNWE", "VBAK-BSTNK", "VBAK-AUGRU", "VBAK-VBTYP", _
        "KUAGV-KUNNR", "KUWEV-KUNNR", "KUNAG-KUNNR", "KUNWE-KUNNR", "AG-KUNNR", _
        "VBAP-MATNR", "VBAP-KWMENG", "VBAP-VRKME", "VBAP-WERKS", "VBAP-PSTYV", _
        "VBAP-CHARG", "VBAP-ETDAT", "VBAP-LPRIO", "VBAP-VSTEL", "VBAP-PRCTR", "VBAP-ARKTX", "VBAP-NETWR", "VBAP-POSNR", _
        "RV45A-MABNR", "VBKD-BSTKD", "VBKD-BSTDK", "VBPA-KUNNR", "VBPA-PARVW" _
    )
End Function

Sub AddContainerPath(ByRef paths, path)
    If Len(path) = 0 Then Exit Sub
    If Not paths.Exists(path) Then paths.Add path, True
End Sub

Sub BuildDynamicContainerPaths(session, usrPath, prog, scr, ByRef paths)
    Dim n, headerScr, bodyPath, subPath, subPart, ctl, anchors, ai
    AddContainerPath paths, usrPath
    If Len(prog) = 0 Then Exit Sub

    If IsNumeric(scr) Then
        For n = CLng(scr) - 3 To CLng(scr) + 3
            headerScr = usrPath & "/subSUBSCREEN_HEADER:" & prog & ":" & CStr(n)
            Set ctl = session.FindById(headerScr)
            If Not ctl Is Nothing Then
                AddContainerPath paths, headerScr
                For subPart = 4698 To 4702
                    subPath = headerScr & "/subPART-SUB:" & prog & ":" & subPart
                    If Not session.FindById(subPath) Is Nothing Then AddContainerPath paths, subPath
                Next
            End If
            bodyPath = usrPath & "/subSUBSCREEN_BODY:" & prog & ":" & CStr(n)
            If Not session.FindById(bodyPath) Is Nothing Then AddContainerPath paths, bodyPath
        Next
    End If

    anchors = Array(4001, 4021, 4301, 4400, 4401, 4900)
    For ai = 0 To UBound(anchors)
        headerScr = usrPath & "/subSUBSCREEN_HEADER:" & prog & ":" & anchors(ai)
        If Not session.FindById(headerScr) Is Nothing Then AddContainerPath paths, headerScr
        bodyPath = usrPath & "/subSUBSCREEN_BODY:" & prog & ":" & anchors(ai)
        If Not session.FindById(bodyPath) Is Nothing Then AddContainerPath paths, bodyPath
    Next

    Set ctl = session.FindById(usrPath & "/tabsTAXI_TABSTRIP_OVERVIEW")
    If Not ctl Is Nothing Then AddContainerPath paths, usrPath & "/tabsTAXI_TABSTRIP_OVERVIEW"
    Set ctl = session.FindById(usrPath & "/tabsTAXI_TABSTRIP")
    If Not ctl Is Nothing Then AddContainerPath paths, usrPath & "/tabsTAXI_TABSTRIP"
End Sub

Sub ProbeFindByIdInContainers(session, ByRef paths, ByRef suffixes, ByRef arr, ByRef count, ByRef seen)
    Dim path, suffix, fieldId, key
    For Each key In paths.Keys
        path = key
        For Each suffix In suffixes.Keys
            TryAppendFieldPath session, path & "/ctxt" & suffix, arr, count, seen
            TryAppendFieldPath session, path & "/txt" & suffix, arr, count, seen
            TryAppendFieldPath session, path & "/cmb" & suffix, arr, count, seen
        Next
    Next
End Sub

Sub ProbeFindByNameFields(session, container, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    If container Is Nothing Then Exit Sub
    Dim names, types, ni, ti, name, typ, coll, i, ctl
    names = Array( _
        "KUAGV-KUNNR", "KUWEV-KUNNR", "KUNAG-KUNNR", "KUNWE-KUNNR", "AG-KUNNR", "KUNNR", _
        "VBAP-MATNR", "MATNR", "RV45A-MABNR", "MABNR", _
        "VBAP-KWMENG", "KWMENG", "VBAP-VRKME", "VRKME", "VBAP-ARKTX", "ARKTX", _
        "VBAK-AUART", "AUART", "VBAK-VKORG", "VKORG", "VBAK-VTWEG", "VTWEG", "VBAK-SPART", "SPART", _
        "S_EBELN", "S_LIFNR", "S_MATNR", "S_WERKS", "S_BUKRS", "S_BEDAT", "S_EKORG", "S_EKGRP", "S_BSART", _
        "EBELN", "LIFNR", "EKORG", "EKGRP", "BSART" _
    )
    types = Array("GuiTextField", "GuiCTextField", "GuiComboBox")
    For ni = 0 To UBound(names)
        name = names(ni)
        For ti = 0 To UBound(types)
            typ = types(ti)
            Err.Clear
            Set ctl = container.FindByName(name, typ)
            If Not ctl Is Nothing Then
                ProcessSapControl session, ctl, arr, count, seen
                        ProcessSapControl session, ctl, arr, count, seen
            End If
            Err.Clear
            Set coll = container.FindAllByName(name, typ)
            If Err.Number = 0 And Not coll Is Nothing Then
                For i = 0 To coll.Count - 1
                    Set ctl = coll.Item(i)
                    ProcessSapControl session, ctl, arr, count, seen
                    If Not ctl Is Nothing Then ProcessSapControl session, ctl, arr, count, seen
                Next
            End If
        Next
    Next
End Sub

Function ExtractSuffixFromControlId(id)
    Dim tail, bracket
    tail = id
    bracket = InStr(tail, "[")
    If bracket > 0 Then tail = Left(tail, bracket - 1)
    tail = Mid(tail, InStrRev(tail, "/") + 1)
    tail = Replace(tail, "ctxt", "", 1, 1, vbTextCompare)
    tail = Replace(tail, "txt", "", 1, 1, vbTextCompare)
    tail = Replace(tail, "cmb", "", 1, 1, vbTextCompare)
    ExtractSuffixFromControlId = tail
End Function

Function ShortSuffixFromFieldName(fieldName)
    Dim dash
    dash = InStrRev(fieldName, "-")
    If dash > 0 Then
        ShortSuffixFromFieldName = Mid(fieldName, dash + 1)
    Else
        ShortSuffixFromFieldName = fieldName
    End If
End Function

Sub AddTitleToMap(ByRef titleMap, fieldName, title)
    Dim shortName
    If Len(fieldName) = 0 Or Len(title) = 0 Then Exit Sub
    If IsGenericColumnTitle(title) Then Exit Sub
    If StrComp(title, fieldName, vbTextCompare) = 0 Then Exit Sub
    If StrComp(title, ShortSuffixFromFieldName(fieldName), vbTextCompare) = 0 Then Exit Sub
    If Not titleMap.Exists(fieldName) Then titleMap.Add fieldName, title
    shortName = ShortSuffixFromFieldName(fieldName)
    If Len(shortName) > 0 And Not titleMap.Exists(shortName) Then titleMap.Add shortName, title
End Sub

Function GetColumnTitleForIndex(tbl, col)
    On Error Resume Next
    Dim title, column
    title = ""

    Err.Clear
    Set column = Nothing
    Set column = tbl.Columns(CLng(col))
    If column Is Nothing Then Set column = tbl.Columns.Item(CLng(col))
    If Not column Is Nothing Then
        title = Trim(column.Title)
        If Len(title) = 0 Then title = Trim(column.Tooltip)
        If Len(title) = 0 Then title = Trim(column.Name)
        If Len(title) = 0 Then title = Trim(column.TechName)
    End If
    If Len(title) > 0 Then
        GetColumnTitleForIndex = Trim(title)
        If IsGenericColumnTitle(GetColumnTitleForIndex) Then GetColumnTitleForIndex = ""
        Exit Function
    End If

    Err.Clear
    title = tbl.GetDisplayedColumnTitle(CLng(col))
    If Len(title) > 0 Then
        GetColumnTitleForIndex = Trim(title)
        If IsGenericColumnTitle(GetColumnTitleForIndex) Then GetColumnTitleForIndex = ""
        Exit Function
    End If

    Err.Clear
    title = tbl.GetColumnTitle(CLng(col))
    If Len(title) > 0 Then
        GetColumnTitleForIndex = Trim(title)
        If IsGenericColumnTitle(GetColumnTitleForIndex) Then GetColumnTitleForIndex = ""
        Exit Function
    End If

    GetColumnTitleForIndex = ""
End Function

Function GetColumnFieldName(tbl, col)
    On Error Resume Next
    Dim fieldName, column
    fieldName = ""
    Err.Clear
    Set column = Nothing
    Set column = tbl.Columns(CLng(col))
    If column Is Nothing Then Set column = tbl.Columns.Item(CLng(col))
    If Not column Is Nothing Then
        fieldName = Trim(column.Name)
        If Len(fieldName) = 0 Then fieldName = Trim(column.TechName)
    End If
    If Len(fieldName) = 0 Then
        Err.Clear
        fieldName = Trim(tbl.GetColumnName(CLng(col)))
    End If
    GetColumnFieldName = fieldName
End Function

Sub BuildTableColumnTitleMap(session, tbl, tblPath, ByRef titleMap)
    On Error Resume Next
    Dim col, maxCol, title, fieldName, suffix, cell, row, c, titleCounts, k
    Set titleMap = CreateObject("Scripting.Dictionary")
    Set titleCounts = CreateObject("Scripting.Dictionary")
    maxCol = 16
    If tbl.ColumnCount > 0 And tbl.ColumnCount < maxCol Then maxCol = tbl.ColumnCount
    If Not tbl.Columns Is Nothing Then
        If tbl.Columns.Count > 0 And tbl.Columns.Count < maxCol Then maxCol = tbl.Columns.Count
    End If

    For col = 0 To maxCol - 1
        fieldName = GetColumnFieldName(tbl, col)
        title = GetColumnTitleForIndex(tbl, col)
        If Len(title) > 0 Then
            If titleCounts.Exists(title) Then
                titleCounts(title) = titleCounts(title) + 1
            Else
                titleCounts.Add title, 1
            End If
        End If
        If Len(fieldName) > 0 And Len(title) > 0 Then AddTitleToMap titleMap, fieldName, title
    Next

    For row = 0 To 1
        For c = 0 To maxCol - 1
            Err.Clear
            Set cell = Nothing
            Set cell = tbl.GetCell(row, c)
            If Err.Number = 0 And Not cell Is Nothing Then
                suffix = ExtractSuffixFromControlId(cell.Id)
                If Len(suffix) > 0 Then
                    If Not titleMap.Exists(suffix) And Not titleMap.Exists(ShortSuffixFromFieldName(suffix)) Then
                        title = GetColumnTitleForIndex(tbl, c)
                        If Len(title) > 0 Then
                            If titleCounts.Exists(title) Then
                                titleCounts(title) = titleCounts(title) + 1
                            Else
                                titleCounts.Add title, 1
                            End If
                            AddTitleToMap titleMap, suffix, title
                        End If
                    End If
                End If
            End If
        Next
    Next

    For Each k In titleCounts.Keys
        If titleCounts(k) > 1 Then RemoveTitleFromMap titleMap, CStr(k)
    Next
End Sub

Sub RemoveTitleFromMap(ByRef titleMap, title)
    Dim keys, i, key
    If titleMap Is Nothing Or Len(title) = 0 Then Exit Sub
    keys = titleMap.Keys
    For i = 0 To UBound(keys)
        key = keys(i)
        If StrComp(titleMap(key), title, vbTextCompare) = 0 Then titleMap.Remove key
    Next
End Sub

Function LookupColumnTitle(ByRef titleMap, fieldId, fieldName)
    Dim suffix, shortName, title
    suffix = ExtractSuffixFromControlId(fieldId)
    If Len(suffix) > 0 And titleMap.Exists(suffix) Then
        LookupColumnTitle = titleMap(suffix)
        Exit Function
    End If
    shortName = ShortSuffixFromFieldName(suffix)
    If Len(shortName) = 0 Then shortName = fieldName
    If Len(shortName) > 0 And titleMap.Exists(shortName) Then
        LookupColumnTitle = titleMap(shortName)
        Exit Function
    End If
    If Len(fieldName) > 0 And titleMap.Exists(fieldName) Then
        LookupColumnTitle = titleMap(fieldName)
        Exit Function
    End If
    LookupColumnTitle = ""
End Function

Sub CaptureTableColumnHeaders(session, tbl, tblPath, ByRef arr, ByRef count, ByRef seen)
    Dim titleMap
    BuildTableColumnTitleMap session, tbl, tblPath, titleMap
End Sub

Sub CaptureTableFields(session, tblPath, tbl, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim row, firstRow, lastRow, rowCount, visibleRows, colCount, maxCol, cell, txt, c, titleMap, colTitle, nm, t, cellLabel
    If tbl Is Nothing Or Len(tblPath) = 0 Then Exit Sub

    rowCount = 0
    visibleRows = 0
    colCount = 0
    rowCount = CLng(tbl.RowCount)
    visibleRows = CLng(tbl.VisibleRowCount)
    colCount = CLng(tbl.ColumnCount)
    If colCount <= 0 Then colCount = CLng(tbl.Columns.Count)
    If colCount <= 0 Then Exit Sub

    BuildTableColumnTitleMap session, tbl, tblPath, titleMap

    firstRow = 0
    Err.Clear
    firstRow = CLng(tbl.FirstVisibleRow)
    If Err.Number <> 0 Then firstRow = 0

    If visibleRows > 0 Then
        lastRow = visibleRows - 1
    ElseIf rowCount > 0 Then
        lastRow = rowCount - 1
    Else
        lastRow = 0
    End If
    If lastRow > 19 Then lastRow = 19

    maxCol = colCount - 1
    If maxCol > 39 Then maxCol = 39

    For row = 0 To lastRow
        For c = 0 To maxCol
            Err.Clear
            Set cell = tbl.GetCell(row, c)
            If Err.Number = 0 And Not cell Is Nothing Then
                txt = ReadFieldValue(cell)
                nm = ResolveControlName(cell)
                cellLabel = ReadFieldLabel(cell, cell.Type)
                colTitle = LookupColumnTitle(titleMap, cell.Id, nm)
                If Len(cellLabel) = 0 Then cellLabel = colTitle
                If Len(cellLabel) = 0 Then cellLabel = GetColumnTitleForIndex(tbl, c)
                If Len(txt) > 0 And Len(cell.Id) > 0 And Not seen.Exists(cell.Id) Then
                    t = cell.Type
                    If Not IsTechnicalType(t) Then
                        seen.Add cell.Id, True
                        AppendControlNode cell, t, cell.Id, txt, nm, arr, count, cellLabel
                    End If
                End If
            End If
        Next
    Next
End Sub

Sub ProbeOverviewFastPaths(session, usrPath, prog, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim headerNums, partNums, partnerSuffixes, hi, pi, fi, hdr, fieldId, tblPaths, ti, tbl, tblPath
    If Len(prog) = 0 Then Exit Sub

    headerNums = Array(4021, 4301, 4020, 4300)
    partNums = Array(4698, 4699, 4700, 4701, 4702)
    partnerSuffixes = Array("KUAGV-KUNNR", "KUWEV-KUNNR", "KUNAG-KUNNR", "KUNWE-KUNNR", "VBPA-KUNNR")

    For hi = 0 To UBound(headerNums)
        hdr = usrPath & "/subSUBSCREEN_HEADER:" & prog & ":" & headerNums(hi)
        For pi = 0 To UBound(partNums)
            For fi = 0 To UBound(partnerSuffixes)
                fieldId = hdr & "/subPART-SUB:" & prog & ":" & partNums(pi) & "/ctxt" & partnerSuffixes(fi)
                TryAppendFieldPath session, fieldId, arr, count, seen
            Next
        Next
        For fi = 0 To UBound(partnerSuffixes)
            TryAppendFieldPath session, hdr & "/ctxt" & partnerSuffixes(fi), arr, count, seen
        Next
    Next

    tblPaths = Array( _
        usrPath & "/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:" & prog & ":4400/subSUBSCREEN_TC:" & prog & ":4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG", _
        usrPath & "/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:" & prog & ":4401/subSUBSCREEN_TC:" & prog & ":4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG", _
        usrPath & "/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:" & prog & ":4400/subSUBSCREEN_TC:" & prog & ":4899/tblSAPMV45ATCTRL_U_ERF_AUFTRAG" _
    )
    For ti = 0 To UBound(tblPaths)
        tblPath = tblPaths(ti)
        Set tbl = session.FindById(tblPath)
        If Not tbl Is Nothing Then
            CaptureTableFields session, tblPath, tbl, arr, count, seen
            Exit For
        End If
    Next
End Sub

Sub DiscoverGuiTableControls(session, usrPath, prog, scr, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim tblPath, tbl, tblName, tabIdx, bodyNums, tcNums, bi, ti, bn, tn
    tblName = "tblSAPMV45ATCTRL_U_ERF_AUFTRAG"
    tabIdx = "\01"
    bodyNums = Array(4400, 4401)
    tcNums = Array(4900, 4899)

    For bi = 0 To UBound(bodyNums)
        bn = bodyNums(bi)
        For ti = 0 To UBound(tcNums)
            tn = tcNums(ti)
            tblPath = usrPath & "/tabsTAXI_TABSTRIP_OVERVIEW/tabpT" & tabIdx & "/ssubSUBSCREEN_BODY:" & prog & ":" & bn & "/subSUBSCREEN_TC:" & prog & ":" & tn & "/" & tblName
            Set tbl = session.FindById(tblPath)
            If Not tbl Is Nothing Then
                CaptureTableFields session, tblPath, tbl, arr, count, seen
                Exit Sub
            End If
        Next
    Next
End Sub

Sub DiscoverNestedOverviewFields(session, usrPath, prog, scr, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim paths, suffixes, candidates, i, usrCtl, filledAfterFast
    ProbeOverviewFastPaths session, usrPath, prog, arr, count, seen
    filledAfterFast = CountFilledFields(arr, count)
    If filledAfterFast >= 3 Then Exit Sub

    Set paths = CreateObject("Scripting.Dictionary")
    Set suffixes = CreateObject("Scripting.Dictionary")

    BuildDynamicContainerPaths session, usrPath, prog, scr, paths

    candidates = GetStandardFieldSuffixes()
    For i = 0 To UBound(candidates)
        suffixes(candidates(i)) = True
    Next

    For i = 0 To count - 1
        AddSuffixFromId arr(i)("id"), suffixes
    Next

    ProbeFindByIdInContainers session, paths, suffixes, arr, count, seen

    Set usrCtl = session.FindById(usrPath)
    If Not usrCtl Is Nothing Then ProbeFindByNameFields session, usrCtl, arr, count, seen

    If CountFilledFields(arr, count) < 3 And Len(prog) > 0 Then
        DiscoverGuiTableControls session, usrPath, prog, scr, arr, count, seen
    End If
End Sub

Function CountFilledFields(arr, count)
    Dim i, txt, n
    n = 0
    For i = 0 To count - 1
        txt = Trim(arr(i)("value"))
        If Len(txt) = 0 Then txt = Trim(arr(i)("text"))
        If Len(txt) > 0 And IsCapturedFieldId(arr(i)("id")) Then n = n + 1
    Next
    CountFilledFields = n
End Function

Function NeedsOverviewDiscovery(session, scr, prog)
    NeedsOverviewDiscovery = False
    If StrComp(prog, "SAPMV45A", vbTextCompare) <> 0 Then Exit Function
    On Error Resume Next
    Dim wndText
    wndText = session.FindById("wnd[0]").Text
    If InStr(1, wndText, "Overview", vbTextCompare) > 0 Then
        NeedsOverviewDiscovery = True
        Exit Function
    End If
    If IsNumeric(scr) Then
        If CLng(scr) >= 4300 Or CLng(scr) = 4000 Then NeedsOverviewDiscovery = True
    End If
End Function

Sub ProbeSubscreenContainers(session, baseUsrPath, prog, scr, ByRef visited, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim patterns, p, path, scrNum, offset, scrTry
    patterns = Array( _
        baseUsrPath & "/subSUBSCREEN_HEADER:" & prog & ":" & scr, _
        baseUsrPath & "/subSUBSCREEN_INITIAL:" & prog & ":" & scr, _
        baseUsrPath & "/subSUBSCREEN_BODY:" & prog & ":" & scr, _
        baseUsrPath & "/subSCREEN_HEADER:" & prog & ":" & scr, _
        baseUsrPath & "/subSCREEN:" & prog & ":" & scr, _
        baseUsrPath & "/tabsTAXI_TABSTRIP", _
        baseUsrPath & "/tabsTAXI_TABSTRIP_OVERVIEW", _
        baseUsrPath & "/tabsTAXI_TABSTRIP_ITEM" _
    )
    For Each p In patterns
        DiscoverByIdBfs session, p, visited, arr, count, seen, 0
    Next
    If IsNumeric(scr) Then
        scrNum = CLng(scr)
        For offset = -3 To 3
            scrTry = CStr(scrNum + offset)
            path = baseUsrPath & "/subSUBSCREEN_HEADER:" & prog & ":" & scrTry
            DiscoverByIdBfs session, path, visited, arr, count, seen, 0
            path = baseUsrPath & "/subSUBSCREEN_BODY:" & prog & ":" & scrTry
            DiscoverByIdBfs session, path, visited, arr, count, seen, 0
        Next
    End If
End Sub

Function ShouldDeepIndexChild(ctl)
    On Error Resume Next
    Dim t, id
    ShouldDeepIndexChild = False
    If ctl Is Nothing Then Exit Function
    t = LCase(ctl.Type)
    id = LCase(ctl.Id)
    If InStr(t, "container") > 0 Or InStr(t, "shell") > 0 Or InStr(t, "custom") > 0 Then
        ShouldDeepIndexChild = True
        Exit Function
    End If
    If InStr(t, "userarea") > 0 Or InStr(t, "table") > 0 Or InStr(t, "grid") > 0 Then
        ShouldDeepIndexChild = True
        Exit Function
    End If
    If InStr(id, "cntl") > 0 Or InStr(id, "grid") > 0 Or InStr(id, "alv") > 0 Or InStr(id, "sub") > 0 Then
        ShouldDeepIndexChild = True
    End If
End Function

Sub DiscoverByDeepIndexScan(session, root, ByRef arr, ByRef count, ByRef seen, maxDepth)
    On Error Resume Next
    If root Is Nothing Then Exit Sub
    If maxDepth <= 0 Then
        ProcessSapControl session, root, arr, count, seen
        Exit Sub
    End If

    Dim d0, d1, d2, i0, i1, i2, id, cid, lt, childCount, max0, max1, max2
    childCount = root.Children.Count
    If childCount <= 0 Then Exit Sub
    max0 = childCount - 1
    If max0 > 150 Then max0 = 150
    max1 = 25
    max2 = 15

    For i0 = 0 To max0
        Err.Clear
        Set d0 = root.Children(i0)
        If Err.Number <> 0 Then Exit For
        ProcessSapControl session, d0, arr, count, seen
        id = d0.Id
        cid = LCase(id)
        lt = LCase(d0.Type)
        If InStr(cid, "cntl") > 0 Or InStr(cid, "grid") > 0 Then ProbeShellPathsForCntl session, id, arr, count, seen
        If InStr(lt, "table") > 0 Or InStr(cid, "/tbl") > 0 Then CaptureTableFields session, id, d0, arr, count, seen
        TryCaptureGridControl session, d0, id, arr, count, seen
        If maxDepth >= 2 And ShouldDeepIndexChild(d0) Then
            For i1 = 0 To max1
                Err.Clear
                Set d1 = d0.Children(i1)
                If Err.Number <> 0 Then Exit For
                ProcessSapControl session, d1, arr, count, seen
                id = d1.Id
                cid = LCase(id)
                lt = LCase(d1.Type)
                If InStr(cid, "cntl") > 0 Or InStr(cid, "grid") > 0 Then ProbeShellPathsForCntl session, id, arr, count, seen
                If InStr(lt, "table") > 0 Then CaptureTableFields session, id, d1, arr, count, seen
                TryCaptureGridControl session, d1, id, arr, count, seen
                If maxDepth >= 3 And ShouldDeepIndexChild(d1) Then
                    For i2 = 0 To max2
                        Err.Clear
                        Set d2 = d1.Children(i2)
                        If Err.Number <> 0 Then Exit For
                        ProcessSapControl session, d2, arr, count, seen
                        TryCaptureGridControl session, d2, d2.Id, arr, count, seen
                    Next
                End If
            Next
        End If
    Next
End Sub

Function ExtractTableFieldSuffix(id)
    Dim pos, tail
    ExtractTableFieldSuffix = ""
    If Len(id) = 0 Then Exit Function
    pos = InStrRev(id, "/")
    If pos > 0 Then
        tail = Mid(id, pos + 1)
    Else
        tail = id
    End If
    tail = Replace(tail, "ctxt", "", 1, 1, vbTextCompare)
    tail = Replace(tail, "lbl", "", 1, 1, vbTextCompare)
    tail = Replace(tail, "txt", "", 1, 1, vbTextCompare)
    tail = Replace(tail, "cmb", "", 1, 1, vbTextCompare)
    If InStr(tail, "-") > 0 And Len(tail) >= 4 Then ExtractTableFieldSuffix = tail
End Function

Sub AddSuffixFromId(id, ByRef suffixes)
    Dim suffix
    suffix = ExtractTableFieldSuffix(id)
    If Len(suffix) > 0 Then suffixes(suffix) = True
End Sub

Sub ProbeFindByIdPairsForSuffixes(session, usrPath, ByRef suffixes, ByRef arr, ByRef count, ByRef seen)
    Dim suffix, fieldId
    For Each suffix In suffixes.Keys
        fieldId = usrPath & "/ctxt" & suffix
        AppendFieldById session, fieldId, arr, count, seen
    Next
End Sub

Sub ProbeFlatUsrFindByIdFields(session, usrPath, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim candidates, i, suffix, fieldId, fld
    candidates = GetStandardFieldSuffixes()
    For i = 0 To UBound(candidates)
        suffix = candidates(i)
        fieldId = usrPath & "/ctxt" & suffix
        Set fld = session.FindById(fieldId)
        If Not fld Is Nothing Then AppendFieldById session, fieldId, arr, count, seen
    Next
End Sub

Function CountEditableFields(arr, count)
    Dim i, id, n
    n = 0
    For i = 0 To count - 1
        id = arr(i)("id")
        If IsEditableFieldId(id) Then n = n + 1
    Next
    CountEditableFields = n
End Function

Function ReadSessionStatusBar(session)
    On Error Resume Next
    Dim sb, txt
    ReadSessionStatusBar = ""
    Set sb = session.FindById("wnd[0]/sbar")
    If sb Is Nothing Then Exit Function
    txt = SafeStr(sb.Text)
    If Len(Trim(txt)) = 0 Then txt = SafeStr(sb.MessageText)
    ReadSessionStatusBar = Trim(txt)
End Function

Function IsSuccessStatusBar(msg)
    Dim s, lower
    IsSuccessStatusBar = False
    s = Trim(msg)
    If Len(s) = 0 Then Exit Function
    lower = LCase(s)
    If InStr(lower, "success") > 0 Then IsSuccessStatusBar = True: Exit Function
    If InStr(lower, "has been saved") > 0 Then IsSuccessStatusBar = True: Exit Function
    If InStr(lower, "has been created") > 0 Then IsSuccessStatusBar = True: Exit Function
    If InStr(lower, "has been posted") > 0 Then IsSuccessStatusBar = True: Exit Function
    If InStr(lower, "has been changed") > 0 Then IsSuccessStatusBar = True: Exit Function
    If InStr(lower, "has been updated") > 0 Then IsSuccessStatusBar = True: Exit Function
    If InStr(lower, "has been released") > 0 Then IsSuccessStatusBar = True: Exit Function
    If InStr(lower, "was saved") > 0 Then IsSuccessStatusBar = True: Exit Function
    If InStr(lower, "was created") > 0 Then IsSuccessStatusBar = True: Exit Function
    If InStr(lower, "was posted") > 0 Then IsSuccessStatusBar = True: Exit Function
End Function

Sub PrepareScreenForDiscovery(session)
    On Error Resume Next
    Dim wnd
    Set wnd = session.FindById("wnd[0]")
    If Not wnd Is Nothing Then wnd.SetFocus
    Set wnd = session.FindById("wnd[0]/usr")
    If Not wnd Is Nothing Then wnd.SetFocus
End Sub

Sub WalkControl(session, ctl, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim i, child, kids, childCount, indexBase, startIndex, endIndex, id, t
    If ctl Is Nothing Then Exit Sub
    If CaptureLimitReached(count) Then Exit Sub

    ProcessSapControl session, ctl, arr, count, seen
    If CaptureLimitReached(count) Then Exit Sub

    id = ctl.Id
    t = LCase(ctl.Type)
    If InStr(t, "table") > 0 Or InStr(LCase(id), "/tbl") > 0 Then
        CaptureTableFields session, id, ctl, arr, count, seen
    End If

    Err.Clear
    Set kids = ctl.Children
    If Err.Number <> 0 Or kids Is Nothing Then
        Err.Clear
        Exit Sub
    End If

    childCount = 0
    childCount = kids.Count
    If Err.Number <> 0 Or childCount <= 0 Then
        Err.Clear
        Exit Sub
    End If

    indexBase = CollectionIndexBase(kids)
    startIndex = indexBase
    endIndex = childCount - 1 + indexBase

    For i = startIndex To endIndex
        If CaptureLimitReached(count) Then Exit For
        Err.Clear
        Set child = ChildAt(ctl, kids, i)

        If Not child Is Nothing Then
            WalkControl session, child, arr, count, seen
        End If
    Next
End Sub

Function ScoreSession(session, winTitle)
    On Error Resume Next
    Dim score, info, wndText, txn
    score = 0
    Set info = session.Info
    If Not info Is Nothing Then
        txn = info.Transaction
        If Len(txn) > 0 Then score = score + 2
    End If
    wndText = session.FindById("wnd[0]").Text
    If Len(wndText) > 0 Then score = score + 2
    If Len(winTitle) > 0 And Len(wndText) > 0 Then
        If InStr(1, winTitle, wndText, vbTextCompare) > 0 Or InStr(1, wndText, winTitle, vbTextCompare) > 0 Then
            score = score + 10
        End If
    End If
    ScoreSession = score
End Function

Function AttachSapSession(application, winTitle)
    On Error Resume Next
    Dim connection, session, best, connCount, sessCount, activeConn, activeSess

    SessionDbgInit()

    connCount = 0
    connCount = application.Children.Count
    SessionDbgSet "connectionCount", connCount

    Set activeConn = application.ActiveConnection
    Set activeSess = application.ActiveSession
    If Not activeConn Is Nothing Then SessionDbgSet "hasActiveConnection", "true"
    If Not activeSess Is Nothing Then
        SessionDbgSet "hasActiveSession", "true"
        SessionDbgSet "activeSessionTxn", activeSess.Info.Transaction
        SessionDbgSet "activeSessionProgram", activeSess.Info.Program
        SessionDbgSet "activeSessionScreen", activeSess.Info.ScreenNumber
    End If

    If Not IsObject(connection) Then
        If connCount > 0 Then
            Set connection = application.Children(0)
            SessionDbgSet "connectionVia", "application.children(0)"
        ElseIf Not activeConn Is Nothing Then
            Set connection = activeConn
            SessionDbgSet "connectionVia", "ActiveConnection"
        End If
    End If

    If IsObject(connection) Then
        sessCount = connection.Children.Count
        SessionDbgSet "sessionCountInConnection", sessCount
    End If

    If Not IsObject(session) Then
        If IsObject(connection) Then
            If connection.Children.Count > 0 Then
                Set session = connection.Children(0)
                SessionDbgSet "sessionVia", "connection.children(0)"
            End If
        End If
    End If

    If IsObject(WScript) Then
        If IsObject(session) Then WScript.ConnectObject session, "on"
        WScript.ConnectObject application, "on"
        SessionDbgSet "wscriptConnectObject", "true"
    End If

    Set best = GetBestSession(application, winTitle)
    If Not best Is Nothing Then
        SessionDbgSet "attachMethod", "GetBestSession"
        Set AttachSapSession = best
    ElseIf Not activeSess Is Nothing Then
        SessionDbgSet "attachMethod", "ActiveSession"
        Set AttachSapSession = activeSess
    ElseIf IsObject(session) Then
        SessionDbgSet "attachMethod", "connection.children(0)"
        Set AttachSapSession = session
    Else
        SessionDbgSet "attachMethod", "none"
        SessionDbgSet "sessionAttached", "false"
        Set AttachSapSession = Nothing
        Exit Function
    End If

    SessionDbgSet "sessionAttached", "true"
    SessionDbgSet "attachedTxn", AttachSapSession.Info.Transaction
    SessionDbgSet "attachedProgram", AttachSapSession.Info.Program
    SessionDbgSet "attachedScreen", AttachSapSession.Info.ScreenNumber
    SessionDbgSet "attachedWindowTitle", AttachSapSession.FindById("wnd[0]").Text
End Function

Function GetBestSession(app, winTitle)
    On Error Resume Next
    Dim session, best, bestScore, score
    Dim connections, conn, ci, sessions, si, s

    Set best = Nothing
    bestScore = -1

    Set session = app.ActiveSession
    If Not session Is Nothing Then
        score = ScoreSession(session, winTitle)
        If score > bestScore Then
            bestScore = score
            Set best = session
        End If
    End If

    Set connections = app.Children
    If connections Is Nothing Then
        Set GetBestSession = best
        Exit Function
    End If

    For ci = 0 To connections.Count - 1
        Set conn = connections.Item(ci)
        Set sessions = conn.Children
        If Not sessions Is Nothing Then
            For si = 0 To sessions.Count - 1
                Set s = sessions.Item(si)
                score = ScoreSession(s, winTitle)
                If score > bestScore Then
                    bestScore = score
                    Set best = s
                End If
            Next
        End If
    Next
    Set GetBestSession = best
End Function

Sub EmitJson(controls, ctrlCount, txn, prog, scr, captureSource, errMsg, sessionConnected)
    Dim json, i, c, typeKey, typeList
    typeList = ""
    If Not gDbgTypeSamples Is Nothing Then
        For Each typeKey In gDbgTypeSamples.Keys
            If Len(typeList) > 0 Then typeList = typeList & ","
            typeList = typeList & """" & JsonEscape(CStr(typeKey)) & """"
        Next
    End If
    json = "{"
    json = json & """timestamp"":""" & JsonEscape(FormatDateTime(Now, 0)) & ""","
    json = json & """transaction"":""" & JsonEscape(txn) & ""","
    json = json & """program"":""" & JsonEscape(prog) & ""","
    json = json & """screen"":""" & JsonEscape(scr) & ""","
    json = json & """captureSource"":""" & JsonEscape(captureSource) & ""","
    json = json & """controls"":["
    For i = 0 To ctrlCount - 1
        If i > 0 Then json = json & ","
        Set c = controls(i)
        json = json & "{"
        json = json & """id"":""" & JsonEscape(SafeStr(c("id"))) & ""","
        json = json & """type"":""" & JsonEscape(SafeStr(c("type"))) & ""","
        json = json & """name"":""" & JsonEscape(SafeStr(c("name"))) & ""","
        json = json & """text"":""" & JsonEscape(SafeStr(c("text"))) & ""","
        json = json & """value"":""" & JsonEscape(SafeStr(c("value"))) & ""","
        json = json & """tooltip"":""" & JsonEscape(SafeStr(c("tooltip"))) & ""","
        json = json & """left"":" & JsonNum(c("left")) & ","
        json = json & """top"":" & JsonNum(c("top")) & ","
        json = json & """width"":" & JsonNum(c("width")) & ","
        json = json & """height"":" & JsonNum(c("height"))
        On Error Resume Next
        If Not IsEmpty(c("label")) Then
            If Len(SafeStr(c("label"))) > 0 Then
                json = json & ",""label"":""" & JsonEscape(SafeStr(c("label"))) & """"
            End If
        End If
        If Not IsEmpty(c("defaultTooltip")) Then
            If Len(SafeStr(c("defaultTooltip"))) > 0 Then
                json = json & ",""defaultTooltip"":""" & JsonEscape(SafeStr(c("defaultTooltip"))) & """"
            End If
        End If
        json = json & "}"
    Next
    json = json & "],"
    json = json & """debug"":{"
    json = json & """bridge"":""vbscript"","
    json = json & """scriptingEngine"":true,"
    json = json & """sessionConnected"":" & LCase(CStr(sessionConnected)) & ","
    json = json & """controlCount"":" & ctrlCount & ","
    json = json & """discoverMs"":" & gDiscoverMs & ","
    json = json & """visitedControls"":" & gDbgVisited & ","
    json = json & """gridCandidates"":" & gDbgGridCandidates & ","
    json = json & """gridCellsCaptured"":" & gDbgGridCells & ","
    json = json & """sampleTypes"":[" & typeList & "]"
    If Len(SessionDbgJson()) > 0 Then
        json = json & ",""sessionAttach"":{" & SessionDbgJson() & "}"
    End If
    json = json & "}"
    If Len(errMsg) > 0 Then json = json & ",""error"":""" & JsonEscape(errMsg) & """"
    json = json & "}"
    WScript.Echo json
End Sub

Sub DiscoverByIndexScan(session, containerId, ByRef visited, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim container, ctl, i, childId, maxScan
    maxScan = 500
    Set container = session.FindById(containerId)
    If container Is Nothing Then Exit Sub

    For i = 0 To maxScan
        Err.Clear
        Set ctl = container.Children(i)
        If Err.Number <> 0 Then Exit For
        If ctl Is Nothing Then Exit For
        childId = ctl.Id
        If Len(childId) > 0 Then
            If Not visited.Exists(childId) Then
                visited.Add childId, True
                ProcessSapControl session, ctl, arr, count, seen
                DiscoverByIdBfs session, childId, visited, arr, count, seen, 0
            End If
        Else
            ProcessSapControl session, ctl, arr, count, seen
        End If
    Next
End Sub

Function NormalizeControlDedupeKey(id)
    Dim tail, bracket, suffix, row, col
    tail = id
    bracket = InStr(tail, "[")
    If bracket > 0 Then
        suffix = Left(tail, bracket - 1)
        row = Mid(tail, bracket + 1)
        col = Mid(row, InStr(row, ",") + 1)
        row = Left(row, InStr(row, ",") - 1)
        col = Left(col, InStr(col, "]") - 1)
        suffix = ExtractSuffixFromControlId(suffix)
        If Len(suffix) > 0 Then
            NormalizeControlDedupeKey = suffix & "[" & row & "," & col & "]"
            Exit Function
        End If
    End If
    NormalizeControlDedupeKey = ExtractSuffixFromControlId(id)
    If Len(NormalizeControlDedupeKey) = 0 Then NormalizeControlDedupeKey = id
End Function

Sub DedupeCapturedControls(ByRef arr, ByRef count)
    Dim keys, i, key, kept(), keptCount
    If count <= 0 Then
        count = 0
        ReDim arr(0)
        Exit Sub
    End If
    Set keys = CreateObject("Scripting.Dictionary")
    keptCount = 0
    ReDim kept(count - 1)
    For i = 0 To count - 1
        key = NormalizeControlDedupeKey(arr(i)("id"))
        If Not keys.Exists(key) Then
            keys.Add key, True
            Set kept(keptCount) = arr(i)
            keptCount = keptCount + 1
        End If
    Next
    count = keptCount
    If keptCount = 0 Then
        ReDim arr(0)
    Else
        ReDim arr(keptCount - 1)
        For i = 0 To keptCount - 1
            Set arr(i) = kept(i)
        Next
    End If
End Sub

Sub DiscoverScreenControls(session, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim usrPath, prog, scr, info, t0, usrCtl, countAfterWalk
    t0 = Timer
    usrPath = "wnd[0]/usr"
    prog = ""
    scr = ""
    Set info = session.Info
    If Not info Is Nothing Then
        prog = info.Program
        scr = info.ScreenNumber
    End If

    PrepareScreenForDiscovery session

    Set usrCtl = session.FindById(usrPath)
    If Not usrCtl Is Nothing Then
        WalkControl session, usrCtl, arr, count, seen
    End If
    countAfterWalk = count

    SessionDbgSet "recursiveUsrWalk", "true"
    SessionDbgSet "recursiveUsrWalkCount", countAfterWalk
    gDiscoverMs = CLng((Timer - t0) * 1000)
    If gDiscoverMs < 0 Then gDiscoverMs = 0
    DedupeCapturedControls arr, count
End Sub

Sub Main()
    Dim SapGuiAuto, application, session, info, seen
    Dim controls(), ctrlCount
    Dim txn, prog, scr, errMsg, captureSource

    gMaxCaptureFields = 10
    If WScript.Arguments.Count > 0 Then
        If IsNumeric(WScript.Arguments(0)) Then
            gMaxCaptureFields = CInt(WScript.Arguments(0))
            If gMaxCaptureFields < 1 Then gMaxCaptureFields = 1
            If gMaxCaptureFields > 50 Then gMaxCaptureFields = 50
        End If
    End If
    SessionDbgSet "maxCaptureFields", gMaxCaptureFields

    ctrlCount = 0
    ReDim controls(0)
    Set seen = CreateObject("Scripting.Dictionary")
    txn = "": prog = "": scr = "": errMsg = "": captureSource = ""

    On Error Resume Next
    SessionDbgInit()

    If Not IsObject(application) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        If SapGuiAuto Is Nothing Then
            SessionDbgSet "sapGuiFound", "false"
            Call EmitJson(controls, 0, txn, prog, scr, "", "SAPGUI COM object not found", False)
            Exit Sub
        End If
        SessionDbgSet "sapGuiFound", "true"
        Set application = SapGuiAuto.GetScriptingEngine
    End If

    If application Is Nothing Then
        SessionDbgSet "scriptingEngineOk", "false"
        Call EmitJson(controls, 0, txn, prog, scr, "", "GetScriptingEngine is NULL - log off SAP and log back in; check sapgui/user_scripting_per_user and S_SCR authorization", False)
        Exit Sub
    End If
    SessionDbgSet "scriptingEngineOk", "true"

    Set session = AttachSapSession(application, "")
    If session Is Nothing Then
        Call EmitJson(controls, 0, txn, prog, scr, "", "No SAP session open - open a transaction before capture", False)
        Exit Sub
    End If

    Set info = session.Info
    If Not info Is Nothing Then
        txn = info.Transaction
        prog = info.Program
        scr = info.ScreenNumber
    End If

    Dim sbarText
    sbarText = ReadSessionStatusBar(session)
    If IsSuccessStatusBar(sbarText) Then
        SessionDbgSet "statusBarOnly", "true"
        SessionDbgSet "statusBar", sbarText
        captureSource = "scripting-vbs-status"
        gDiscoverMs = 0
    Else
        DiscoverScreenControls session, controls, ctrlCount, seen
        captureSource = "scripting-vbs"
        If ctrlCount = 0 Then errMsg = "Scripting connected but no fields found on screen"
    End If

    Call EmitJson(controls, ctrlCount, txn, prog, scr, captureSource, errMsg, True)
End Sub

Call Main()
