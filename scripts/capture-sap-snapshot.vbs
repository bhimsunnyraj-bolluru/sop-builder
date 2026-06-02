' SAP GUI Scripting bridge — VBScript late binding (PowerShell TYPE_E_CANTLOADLIBRARY workaround).
'
' Generic, transaction-agnostic capture: walks the active SAP window's control
' tree from <window>/usr, reads every value-bearing field (text/combo/check/radio),
' reads ALV grids and table controls as cells, and prunes control-host subtrees
' (ALV/tree/HTML shells) once read. No hardcoded t-code paths or field names.
'
' Bounded by depth / breadth / visited-node / wall-clock caps so it always returns
' within the host timeout instead of hanging on control-heavy screens (ME21N, etc.).
'
' Usage: cscript //Nologo capture-sap-snapshot.vbs [maxFields]
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

' --- Generic safety caps (NOT transaction-specific) -------------------------
Dim gMaxDepth
gMaxDepth = 24
Dim gMaxChildrenPerNode
gMaxChildrenPerNode = 400
Dim gMaxVisited
gMaxVisited = 6000
Dim gBudgetMs
gBudgetMs = 7000
Dim gStartTimer
gStartTimer = 0
Dim gBudgetHit
gBudgetHit = False

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

' Wall-clock guard so a control-heavy screen returns partial results instead of
' being killed by the host timeout. Handles the midnight Timer rollover.
Function BudgetExceeded()
    Dim el
    el = (Timer - gStartTimer) * 1000
    If el < 0 Then el = el + 86400000
    BudgetExceeded = (el > gBudgetMs)
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

' Value-aware read: checkboxes/radios report selection state, everything else
' falls back to the generic text reader.
Function ReadControlValue(ctl, lt)
    On Error Resume Next
    If InStr(lt, "checkbox") > 0 Or InStr(lt, "radiobutton") > 0 Then
        If ctl.Selected Then
            ReadControlValue = "X"
        Else
            ReadControlValue = ""
        End If
        Exit Function
    End If
    ReadControlValue = ReadFieldValue(ctl)
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

Sub DbgNoteType(t)
    On Error Resume Next
    If gDbgTypeSamples Is Nothing Then Set gDbgTypeSamples = CreateObject("Scripting.Dictionary")
    If Len(t) > 0 Then
        If Not gDbgTypeSamples.Exists(t) Then
            If gDbgTypeSamples.Count < 25 Then gDbgTypeSamples.Add t, True
        End If
    End If
End Sub

' Geometry (left/top/width/height) is not consumed downstream, so it is emitted
' as 0 rather than paying 4–8 cross-process COM reads per field.
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
    node.Add "tooltip", SafeStr(ctl.Tooltip)
    node.Add "left", 0
    node.Add "top", 0
    node.Add "width", 0
    node.Add "height", 0
    If Len(fieldLabel) > 0 Then node.Add "label", fieldLabel
    If InStr(LCase(t), "ctextfield") > 0 Then
        rawTip = Trim(SafeStr(ctl.DefaultTooltip))
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
    Dim kids, i, child, gv, lt
    ' ALV grids sit only a few levels under their host; a shallow cap avoids an
    ' expensive deep scan (with throwing GetCellValue/GetGridView calls) on
    ' non-grid shells such as trees, HTML viewers and toolbars.
    If root Is Nothing Or depth > 6 Then Exit Sub
    If Not gridOut Is Nothing Then Exit Sub

    If GridLooksReadable(root) Then
        gDbgGridCandidates = gDbgGridCandidates + 1
        Set gridOut = root
        Exit Sub
    End If

    ' GetGridView() exists only on shell controls and raises (slow) on others.
    lt = LCase(root.Type)
    If InStr(lt, "shell") > 0 Then
        Err.Clear
        Set gv = root.GetGridView()
        If Not gv Is Nothing Then
            If GridLooksReadable(gv) Then
                gDbgGridCandidates = gDbgGridCandidates + 1
                Set gridOut = gv
                Exit Sub
            End If
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
        If CaptureLimitReached(count) Then Exit Sub
        If BudgetExceeded() Then gBudgetHit = True : Exit Sub
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

' True for control-host nodes (ALV/tree/HTML/picture shells, custom containers)
' whose internal subtree should be read as a grid then pruned. Layout splitters
' are excluded so their real content is still walked.
Function ShouldProbeAsGrid(lt)
    ShouldProbeAsGrid = False
    If InStr(lt, "splitter") > 0 Then Exit Function
    If InStr(lt, "gridview") > 0 Then ShouldProbeAsGrid = True : Exit Function
    If InStr(lt, "containershell") > 0 Then ShouldProbeAsGrid = True : Exit Function
    If InStr(lt, "customcontrol") > 0 Then ShouldProbeAsGrid = True : Exit Function
    If InStr(lt, "shell") > 0 Then ShouldProbeAsGrid = True : Exit Function
End Function

' --- Generic table-control (GuiTableControl) reading ------------------------

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

Sub RemoveTitleFromMap(ByRef titleMap, title)
    Dim keys, i, key
    If titleMap Is Nothing Or Len(title) = 0 Then Exit Sub
    keys = titleMap.Keys
    For i = 0 To UBound(keys)
        key = keys(i)
        If StrComp(titleMap(key), title, vbTextCompare) = 0 Then titleMap.Remove key
    Next
End Sub

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

Function LookupColumnTitle(ByRef titleMap, fieldId, fieldName)
    Dim suffix, shortName
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
        If CaptureLimitReached(count) Then Exit Sub
        If BudgetExceeded() Then gBudgetHit = True : Exit Sub
        For c = 0 To maxCol
            Err.Clear
            Set cell = tbl.GetCell(row, c)
            If Err.Number = 0 And Not cell Is Nothing Then
                ' Read the value FIRST and skip empty cells before paying for
                ' name/label/column-title lookups (those cost several COM calls
                ' each; ME21N-style tables are wide and mostly empty).
                txt = ReadFieldValue(cell)
                If Len(txt) > 0 And Len(cell.Id) > 0 And Not seen.Exists(cell.Id) Then
                    t = cell.Type
                    If Not IsTechnicalType(t) Then
                        nm = ResolveControlName(cell)
                        cellLabel = ReadFieldLabel(cell, t)
                        colTitle = LookupColumnTitle(titleMap, cell.Id, nm)
                        If Len(cellLabel) = 0 Then cellLabel = colTitle
                        If Len(cellLabel) = 0 Then cellLabel = GetColumnTitleForIndex(tbl, c)
                        seen.Add cell.Id, True
                        AppendControlNode cell, t, cell.Id, txt, nm, arr, count, cellLabel
                    End If
                End If
            End If
            If CaptureLimitReached(count) Then Exit Sub
            If BudgetExceeded() Then gBudgetHit = True : Exit Sub
        Next
    Next
End Sub

' --- Status bar -------------------------------------------------------------

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

' --- Generic capture core ---------------------------------------------------

Sub CaptureValueField(ctl, t, lt, id, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim txt, nm
    If Len(id) = 0 Then Exit Sub
    If seen.Exists(id) Then Exit Sub
    If IsTechnicalType(t) Then Exit Sub
    If Not IsEditableFieldType(t) Then Exit Sub
    txt = ReadControlValue(ctl, lt)
    If Len(Trim(txt)) = 0 Then Exit Sub
    nm = ResolveControlName(ctl)
    seen.Add id, True
    AppendControlNode ctl, t, id, txt, nm, arr, count, ""
End Sub

' Bounded, generic, depth-first walk. Reads value fields; reads tables/ALV grids
' as cells and prunes their internals; recurses everything else. Stops on any of
' the field/depth/breadth/visited/wall-clock caps.
Sub WalkControl(session, ctl, ByRef arr, ByRef count, ByRef seen, depth)
    On Error Resume Next
    If ctl Is Nothing Then Exit Sub
    If CaptureLimitReached(count) Then Exit Sub
    If depth > gMaxDepth Then Exit Sub
    If gDbgVisited > gMaxVisited Then Exit Sub
    If BudgetExceeded() Then gBudgetHit = True : Exit Sub

    Dim t, lt, id, kids, cc, i, child
    t = ctl.Type
    lt = LCase(t)
    id = ctl.Id
    gDbgVisited = gDbgVisited + 1
    DbgNoteType t

    ' Old-style table control: read cells, then prune (cells are its children).
    If InStr(lt, "table") > 0 Or InStr(LCase(id), "/tbl") > 0 Then
        CaptureTableFields session, id, ctl, arr, count, seen
        Exit Sub
    End If

    ' ALV / control host (GuiShell/ContainerShell/CustomControl/GridView): probe
    ' for a readable grid ONCE, then prune. These host ALV/tree/HTML/picture
    ' controls — never dynpro input fields — so re-walking their internals only
    ' repeats the (expensive) grid descent without finding capturable fields.
    If ShouldProbeAsGrid(lt) Then
        TryCaptureGridControl session, ctl, id, arr, count, seen
        Exit Sub
    End If

    CaptureValueField ctl, t, lt, id, arr, count, seen

    Set kids = ctl.Children
    If Err.Number <> 0 Or kids Is Nothing Then Exit Sub
    cc = kids.Count
    If Err.Number <> 0 Or cc <= 0 Then Exit Sub
    If cc > gMaxChildrenPerNode Then cc = gMaxChildrenPerNode

    ' SAP GuiComponentCollection is 0-based; one Item() call per child (no
    ' index-base probing or multi-fallback) keeps COM round-trips minimal.
    For i = 0 To cc - 1
        If CaptureLimitReached(count) Then Exit For
        If BudgetExceeded() Then gBudgetHit = True : Exit For
        Err.Clear
        Set child = Nothing
        Set child = kids.Item(CLng(i))
        If Err.Number = 0 Then
            If Not child Is Nothing Then WalkControl session, child, arr, count, seen, depth + 1
        End If
    Next
End Sub

' --- Session attach ---------------------------------------------------------

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

' --- Dedupe + emit ----------------------------------------------------------

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
    json = json & """budgetExceeded"":" & LCase(CStr(gBudgetHit)) & ","
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

' Walk the ACTIVE window's user area (handles modal popups generically), with all
' bounds enforced by WalkControl. No transaction-specific paths.
Sub DiscoverScreenControls(session, ByRef arr, ByRef count, ByRef seen)
    On Error Resume Next
    Dim usrPath, basePath, activeWnd, prog, scr, info, t0, usrCtl
    t0 = Timer
    gStartTimer = Timer
    gBudgetHit = False
    prog = ""
    scr = ""
    Set info = session.Info
    If Not info Is Nothing Then
        prog = info.Program
        scr = info.ScreenNumber
    End If

    PrepareScreenForDiscovery session

    basePath = "wnd[0]"
    Set activeWnd = session.ActiveWindow
    If Not activeWnd Is Nothing Then
        If Len(activeWnd.Id) > 0 Then basePath = activeWnd.Id
    End If
    usrPath = basePath & "/usr"

    Set usrCtl = session.FindById(usrPath)
    If usrCtl Is Nothing Then
        usrPath = "wnd[0]/usr"
        Set usrCtl = session.FindById(usrPath)
    End If
    If Not usrCtl Is Nothing Then WalkControl session, usrCtl, arr, count, seen, 0

    SessionDbgSet "discoverWindow", basePath
    SessionDbgSet "budgetExceeded", LCase(CStr(gBudgetHit))
    SessionDbgSet "recursiveUsrWalkCount", count
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
