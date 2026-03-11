Attribute VB_Name = "modMeta"
Option Explicit

Private Const METADATA_DIR As String = ".metadata"

Public Sub EnsureMetadataFolder()
    Dim folder As String
    folder = GetMetadataFolder()
    modEncodingIO.EnsureFolderExists folder
End Sub

Public Function BuildMetaPath(ByVal csvAbsPath As String) As String
    Dim folder As String
    folder = GetMetadataFolder()

    Dim baseName As String
    baseName = GetFileName(csvAbsPath)

    Dim safeName As String
    safeName = MakeFileNameSafe(baseName)

    Dim h As String
    h = modHash.Hex8(modHash.Crc32String(csvAbsPath))

    BuildMetaPath = JoinPath(folder, h & "_" & safeName & ".json")
End Function

Private Function GetMetadataFolder() As String
    GetMetadataFolder = JoinPath(ThisWorkbook.path, METADATA_DIR)
End Function

Private Function JoinPath(ByVal dirPath As String, ByVal rel As String) As String
    If Right$(dirPath, 1) = "\" Or Right$(dirPath, 1) = "/" Then
        JoinPath = dirPath & rel
    Else
        JoinPath = dirPath & "\" & rel
    End If
End Function

Private Function GetFileName(ByVal fullPath As String) As String
    Dim i As Long
    i = InStrRev(fullPath, "\")
    If i = 0 Then
        GetFileName = fullPath
    Else
        GetFileName = Mid$(fullPath, i + 1)
    End If
End Function

Private Function MakeFileNameSafe(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("<", ">", ":", """", "/", "\", "|", "?", "*")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, CStr(bad(i)), "_")
    Next i
    MakeFileNameSafe = s
End Function

' =====================
' Metadata (JSON) read/apply/write
' =====================

Public Sub ApplyMetaConfig(ByVal json As String, ByVal ws As Worksheet, ByVal tableName As String)
    ' Column widths
    Dim widths As Variant
    widths = ExtractNumberArray(json, "column_widths")
    If Not IsEmpty(widths) Then
        Dim i As Long
        For i = LBound(widths) To UBound(widths)
            ws.Columns(i + 1).ColumnWidth = CDbl(widths(i))
        Next i
    End If

    ' Hidden columns
    Dim hiddenCols As Variant
    hiddenCols = ExtractIntArray(json, "hidden_columns")
    If Not IsEmpty(hiddenCols) Then
        Dim j As Long
        For j = LBound(hiddenCols) To UBound(hiddenCols)
            Dim colIx As Long
            colIx = CLng(hiddenCols(j))
            If colIx >= 1 Then ws.Columns(colIx).Hidden = True
        Next j
    End If

    ' Table style
    Dim styleName As String
    styleName = ExtractStringValue(json, "table_style")
    If Len(styleName) > 0 Then
        On Error Resume Next
        ws.ListObjects(tableName).TableStyle = styleName
        On Error GoTo 0
    End If

    ' Freeze panes (row/col)
    Dim fr As Long, fc As Long
    fr = ExtractIntValue(json, "freeze_row", 2)
    fc = ExtractIntValue(json, "freeze_col", 1)

    ' Word wrap (columns)
    Dim wraps As Variant
    wraps = ExtractBoolArray(json, "wrap_columns")
    ' Also get the raw array text to detect "null" entries
    Dim wrapRaw() As String
    Dim wrapInner As String
    wrapInner = Replace$(Replace$(Replace$(ExtractArrayInner(json, "wrap_columns"), vbCr, ""), vbLf, ""), " ", "")
    If Len(wrapInner) > 0 Then
        wrapRaw = Split(wrapInner, ",")
    End If

    If Not IsEmpty(wraps) Then
        Dim k As Long
        For k = LBound(wraps) To UBound(wraps)
            If k <= UBound(wrapRaw) Then
                If wrapRaw(k) <> "null" Then
                    ws.Columns(k + 1).WrapText = CBool(wraps(k))
                End If
                ' null = skip, let cell overrides handle it
            End If
        Next k
    End If

    ' Word wrap (cell overrides)
    Dim cellWraps As Object
    Set cellWraps = ExtractStringDict(json, "wrap_cells")
    If Not cellWraps Is Nothing Then
        Dim addr As Variant
        For Each addr In cellWraps.Keys
            ws.Range(CStr(addr)).WrapText = (cellWraps(addr) = "true")
        Next addr
    End If

    ws.Activate
    ws.Cells(fr, fc).Select
    ActiveWindow.FreezePanes = True
End Sub

Public Function BuildMetaJson(ByVal ws As Worksheet, ByVal tableName As String, ByVal csvPathAbs As String, ByVal delim As String, ByVal enc As String) As String
    Dim used As Range
    Set used = ws.UsedRange

    Dim nCols As Long
    If used Is Nothing Then
        nCols = 0
    Else
        nCols = used.Columns.Count
    End If

    ' widths
    Dim i As Long
    Dim wParts() As String
    If nCols > 0 Then
        ReDim wParts(0 To nCols - 1)
        For i = 1 To nCols
            wParts(i - 1) = InvariantNumber(ws.Columns(i).ColumnWidth)
        Next i
    Else
        ReDim wParts(0 To -1)
    End If

    ' hidden columns
    Dim hParts() As String
    Dim hCount As Long: hCount = 0
    ReDim hParts(0 To IIf(nCols > 0, nCols - 1, 0))

    If nCols > 0 Then
        For i = 1 To nCols
            If ws.Columns(i).Hidden Then
                hParts(hCount) = CStr(i)
                hCount = hCount + 1
            End If
        Next i
    End If
    If hCount > 0 Then
        ReDim Preserve hParts(0 To hCount - 1)
    Else
        Erase hParts
    End If

    ' resolve the window that is displaying this workbook (best-effort) and read freeze state from there.
    Dim freezeRow As Long, freezeCol As Long
    freezeRow = 2
    freezeCol = 1
    GetFreezeFromWorkbookWindow ws, freezeRow, freezeCol

    ' table style
    Dim styleName As String
    styleName = ""
    On Error Resume Next
    styleName = ws.ListObjects(tableName).TableStyle
    On Error GoTo 0

    ' wrap settings (per column)
    Dim wrapParts() As String
    If nCols > 0 Then
        ReDim wrapParts(0 To nCols - 1)
        For i = 1 To nCols
            Dim cw As Variant
            cw = ws.Columns(i).ColumnWidth
            Dim colWrapState As Variant
            colWrapState = ws.Columns(i).WrapText
            If IsNull(colWrapState) Then
                wrapParts(i - 1) = "null"   ' mixed - let cell overrides handle it
            ElseIf CBool(colWrapState) Then
                wrapParts(i - 1) = "true"
            Else
                wrapParts(i - 1) = "false"
            End If
        Next i
    Else
        ReDim wrapParts(0 To -1)
    End If

    ' wrap_cells: only cells that differ from their column default
    Dim cellWrapParts As String
    cellWrapParts = "{"
    Dim firstEntry As Boolean
    firstEntry = True
    
    Dim cell As Range
    Dim colWrap As Variant
    Dim cellWrap As Variant
    
    For Each cell In ws.UsedRange
        cellWrap = cell.WrapText
        If Not IsNull(cellWrap) Then
            Dim colDefault As String
            colDefault = wrapParts(cell.Column - 1)  ' "true", "false", or "null"
            Dim cellBool As Boolean
            cellBool = CBool(cellWrap)
            Dim differs As Boolean
            If colDefault = "null" Then
                differs = True   ' column is mixed, record all cells
            ElseIf colDefault = "true" Then
                differs = Not cellBool
            Else
                differs = cellBool
            End If
            If differs Then
                If Not firstEntry Then cellWrapParts = cellWrapParts & ", "
                cellWrapParts = cellWrapParts & """" & cell.Address(False, False) & """: " & _
                                IIf(cellBool, "true", "false")
                firstEntry = False
            End If
        End If
    Next cell
    cellWrapParts = cellWrapParts & "}"

    Dim json As String
    json = "{" & vbCrLf & _
           "  ""csv_path"": " & JsonString(csvPathAbs) & "," & vbCrLf & _
           "  ""delimiter"": " & JsonString(delim) & "," & vbCrLf & _
           "  ""encoding"": " & JsonString(enc) & "," & vbCrLf & _
           "  ""table_style"": " & JsonString(styleName) & "," & vbCrLf & _
           "  ""freeze_row"": " & CStr(freezeRow) & "," & vbCrLf & _
           "  ""freeze_col"": " & CStr(freezeCol) & "," & vbCrLf & _
           "  ""column_widths"": [" & Join(wParts, ",") & "]," & vbCrLf & _
           "  ""hidden_columns"": [" & Join(hParts, ",") & "]," & vbCrLf & _
           "  ""wrap_columns"": [" & Join(wrapParts, ",") & "]," & vbCrLf & _
           "  ""wrap_cells"": " & cellWrapParts & vbCrLf & _
           "}" & vbCrLf

    BuildMetaJson = json
End Function

Private Sub GetFreezeFromWorkbookWindow(ByVal ws As Worksheet, ByRef freezeRow As Long, ByRef freezeCol As Long)
    On Error GoTo done

    Dim win As Window
    Set win = Nothing

    ' Prefer a window belonging to this workbook.
    If ws.Parent.Windows.Count > 0 Then
        Set win = ws.Parent.Windows(1)
    End If

    If win Is Nothing Then GoTo done

    If win.FreezePanes Then
        freezeRow = win.SplitRow + 1
        freezeCol = win.SplitColumn + 1
    End If

done:
End Sub

Private Function JsonString(ByVal s As String) As String
    s = Replace$(s, "\", "\\")
    s = Replace$(s, """", "\""")
    s = Replace$(s, vbCr, "\r")
    s = Replace$(s, vbLf, "\n")
    JsonString = """" & s & """"
End Function

Private Function InvariantNumber(ByVal v As Double) As String
    Dim s As String
    s = CStr(v)
    s = Replace$(s, ",", ".")
    InvariantNumber = s
End Function

' =====================
' Minimal JSON extractors (for our own generated JSON)
' =====================

Public Function ExtractStringValue(ByVal json As String, ByVal key As String) As String
    On Error GoTo fail
    Dim p As Long, c As Long, q As Long
    p = InStr(1, json, """" & key & """", vbTextCompare)
    If p = 0 Then GoTo fail
    c = InStr(p, json, ":", vbBinaryCompare)
    If c = 0 Then GoTo fail
    q = InStr(c, json, """", vbBinaryCompare)
    If q = 0 Then GoTo fail

    Dim r As Long
    r = q + 1
    Dim out As String: out = ""
    Do While r <= Len(json)
        Dim ch As String
        ch = Mid$(json, r, 1)
        If ch = "\" Then
            Dim nxt As String
            nxt = Mid$(json, r + 1, 1)
            If nxt = """" Or nxt = "\" Then
                out = out & nxt
                r = r + 2
            ElseIf nxt = "n" Then
                out = out & vbLf: r = r + 2
            ElseIf nxt = "r" Then
                out = out & vbCr: r = r + 2
            Else
                r = r + 2
            End If
        ElseIf ch = """" Then
            Exit Do
        Else
            out = out & ch
            r = r + 1
        End If
    Loop

    ExtractStringValue = out
    Exit Function
fail:
    ExtractStringValue = ""
End Function

Private Function ExtractIntValue(ByVal json As String, ByVal key As String, ByVal defaultValue As Long) As Long
    On Error GoTo fail
    Dim p As Long, c As Long
    p = InStr(1, json, """" & key & """", vbTextCompare)
    If p = 0 Then GoTo fail
    c = InStr(p, json, ":", vbBinaryCompare)
    If c = 0 Then GoTo fail

    Dim s As String
    s = ReadNumberToken(Mid$(json, c + 1))
    If Len(s) = 0 Then GoTo fail

    ExtractIntValue = CLng(s)
    Exit Function
fail:
    ExtractIntValue = defaultValue
End Function

Private Function ReadNumberToken(ByVal s As String) As String
    s = Trim$(s)
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr("0123456789-.", ch) = 0 Then Exit For
    Next i
    ReadNumberToken = Left$(s, i - 1)
End Function

Private Function ExtractNumberArray(ByVal json As String, ByVal key As String) As Variant
    On Error GoTo fail
    Dim inner As String
    inner = ExtractArrayInner(json, key)
    If Len(inner) = 0 Then GoTo fail

    inner = Replace$(inner, vbCr, "")
    inner = Replace$(inner, vbLf, "")
    inner = Replace$(inner, " ", "")
    If Len(inner) = 0 Then GoTo fail

    Dim tokens() As String
    tokens = Split(inner, ",")

    Dim out() As Variant
    Dim i As Long
    ReDim out(0 To UBound(tokens))
    For i = 0 To UBound(tokens)
        out(i) = CDbl(tokens(i))
    Next i

    ExtractNumberArray = out
    Exit Function
fail:
    ExtractNumberArray = Empty
End Function

Private Function ExtractIntArray(ByVal json As String, ByVal key As String) As Variant
    On Error GoTo fail
    Dim inner As String
    inner = ExtractArrayInner(json, key)
    If Len(inner) = 0 Then GoTo fail

    inner = Replace$(inner, vbCr, "")
    inner = Replace$(inner, vbLf, "")
    inner = Replace$(inner, " ", "")
    If Len(inner) = 0 Then GoTo fail

    Dim tokens() As String
    tokens = Split(inner, ",")

    Dim out() As Variant
    Dim i As Long
    ReDim out(0 To UBound(tokens))
    For i = 0 To UBound(tokens)
        out(i) = CLng(tokens(i))
    Next i

    ExtractIntArray = out
    Exit Function
fail:
    ExtractIntArray = Empty
End Function

Private Function ExtractArrayInner(ByVal json As String, ByVal key As String) As String
    Dim p As Long, lb As Long, rb As Long
    p = InStr(1, json, """" & key & """", vbTextCompare)
    If p = 0 Then
        ExtractArrayInner = ""
        Exit Function
    End If
    lb = InStr(p, json, "[", vbBinaryCompare)
    If lb = 0 Then
        ExtractArrayInner = ""
        Exit Function
    End If
    rb = InStr(lb, json, "]", vbBinaryCompare)
    If rb = 0 Then
        ExtractArrayInner = ""
        Exit Function
    End If
    ExtractArrayInner = Mid$(json, lb + 1, rb - lb - 1)
End Function

Private Function ExtractBoolArray(ByVal json As String, ByVal key As String) As Variant
    On Error GoTo fail
    Dim inner As String
    inner = ExtractArrayInner(json, key)
    If Len(inner) = 0 Then GoTo fail

    inner = Replace$(inner, vbCr, "")
    inner = Replace$(inner, vbLf, "")
    inner = Replace$(inner, " ", "")
    If Len(inner) = 0 Then GoTo fail

    Dim tokens() As String
    tokens = Split(inner, ",")

    Dim out() As Variant
    Dim i As Long
    ReDim out(0 To UBound(tokens))
    For i = 0 To UBound(tokens)
        Dim t As String
        t = LCase$(tokens(i))
        out(i) = (t = "true" Or t = "1")
    Next i

    ExtractBoolArray = out
    Exit Function
fail:
    ExtractBoolArray = Empty
End Function


Private Function ExtractStringDict(ByVal json As String, ByVal key As String) As Object
    Dim p As Long
    p = InStr(1, json, """" & key & """", vbTextCompare)
    If p = 0 Then Set ExtractStringDict = Nothing: Exit Function

    Dim lb As Long, rb As Long
    lb = InStr(p, json, "{", vbBinaryCompare)
    If lb = 0 Then Set ExtractStringDict = Nothing: Exit Function

    ' Find matching closing brace
    Dim depth As Long: depth = 1
    Dim i As Long
    For i = lb + 1 To Len(json)
        Dim ch As String
        ch = Mid$(json, i, 1)
        If ch = "{" Then depth = depth + 1
        If ch = "}" Then
            depth = depth - 1
            If depth = 0 Then rb = i: Exit For
        End If
    Next i
    If rb = 0 Then Set ExtractStringDict = Nothing: Exit Function

    Dim inner As String
    inner = Mid$(json, lb + 1, rb - lb - 1)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' Parse "KEY": VALUE pairs
    Dim pos As Long: pos = 1
    Do While pos <= Len(inner)
        Dim ks As Long
        ks = InStr(pos, inner, """", vbBinaryCompare)
        If ks = 0 Then Exit Do
        Dim ke As Long
        ke = InStr(ks + 1, inner, """", vbBinaryCompare)
        If ke = 0 Then Exit Do
        Dim entryKey As String
        entryKey = Mid$(inner, ks + 1, ke - ks - 1)

        Dim col As Long
        col = InStr(ke + 1, inner, ":", vbBinaryCompare)
        If col = 0 Then Exit Do

        ' Find next comma or end of inner
        Dim ve As Long
        ve = InStr(col + 1, inner, ",")
        Dim valStr As String
        If ve > 0 Then
            valStr = Trim$(Mid$(inner, col + 1, ve - col - 1))
            pos = ve + 1
        Else
            valStr = Trim$(Mid$(inner, col + 1))
            Exit Do
        End If

        dict(entryKey) = LCase$(valStr)
    Loop

    Set ExtractStringDict = dict
End Function
