Attribute VB_Name = "modCsvEditor"
Option Explicit

' ============================
' Session-only state (not saved)
' ============================
Private gCsvPathAbs As String
Private gCsvPathOriginalArg As String
Private gDelim As String

' Encoding name used for read/write:
'   "utf-8", "windows-1252", "utf-16le", "utf-16be"
Private gEncoding As String

Private gMetaPath As String

' Sheet names
Private Const SHEET_DATA As String = "Data"
Private Const TABLE_NAME As String = "Table1"

' Metadata folder relative to workbook
Private Const METADATA_DIR As String = ".metadata"

Public Function CurrentCsvPathAbs() As String
    CurrentCsvPathAbs = gCsvPathAbs
End Function


' =====================
' Public entry points
' =====================

Public Sub ReloadFromEnv()
    Dim pArg As String
    pArg = GetEnv("EXCEL_CSV_PATH")
    If Len(pArg) = 0 Then
        If Not ThisWorkbook.DevModeEnabled Then
            MsgBox "Required environment variable EXCEL_CSV_PATH is not set." & vbCrLf & _
                   "Launch via the provided .cmd.", vbCritical
        End If
        Exit Sub
    End If

    gCsvPathOriginalArg = pArg
    gCsvPathAbs = ResolveCsvPath(pArg)

    If Len(gCsvPathAbs) = 0 Then
        MsgBox "Could not resolve CSV path from EXCEL_CSV_PATH: " & pArg, vbCritical
        Exit Sub
    End If

    If Dir$(gCsvPathAbs) = "" Then
        MsgBox "CSV file not found: " & gCsvPathAbs, vbCritical
        Exit Sub
    End If

    gMetaPath = BuildMetaPath(gCsvPathAbs)

    ' Load meta first (may set delimiter/encoding/style settings)
    Dim metaJson As String
    metaJson = ReadTextFileBestEffort(gMetaPath, "utf-8")
    Dim metaExists As Boolean
    metaExists = (Len(metaJson) > 0)

    ' Delimiter selection:
    ' 1) EXCEL_CSV_DELIM (if set)
    ' 2) metadata (if present)
    ' 3) auto-detect (comma vs semicolon)
    ' 4) default comma
    Dim envDelim As String
    envDelim = GetEnv("EXCEL_CSV_DELIM")
    If Len(envDelim) > 0 Then
        gDelim = Left$(envDelim, 1)
    ElseIf metaExists Then
        Dim md As String
        md = ExtractStringValue(metaJson, "delimiter")
        If Len(md) > 0 Then gDelim = Left$(md, 1)
    End If
    If Len(gDelim) = 0 Then
        gDelim = DetectDelimiterBestEffort(gCsvPathAbs, ",")
    End If

    ' Encoding selection:
    ' 1) metadata (if present) to preserve previous detected encoding
    ' 2) detect from file
    gEncoding = ""
    If metaExists Then
        gEncoding = ExtractStringValue(metaJson, "encoding")
    End If
    If Len(gEncoding) = 0 Then
        gEncoding = DetectEncoding(gCsvPathAbs)
    End If

    ' Load CSV to Data sheet
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets(SHEET_DATA)

    LoadCsvIntoSheet gCsvPathAbs, gDelim, gEncoding, wsData

    ' Apply meta formatting/settings
    If metaExists Then
        ApplyMetaConfig metaJson, wsData
    Else
        ApplyDefaultView wsData
    End If
    
    modControls.UpdateCsvPathLabel
End Sub


Public Sub ExportToOriginalCsv(Optional ByVal showConfirmation As Boolean = True)
    If Len(gCsvPathAbs) = 0 Then
        MsgBox "No CSV loaded in this session. Use ReloadFromEnv first.", vbExclamation
        Exit Sub
    End If

    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets(SHEET_DATA)

    ExportSheetToCsv gCsvPathAbs, gDelim, gEncoding, wsData

    ' Always write metadata after export (captures current view settings)
    Dim metaJson As String
    metaJson = BuildMetaJson(wsData)
    EnsureFolderExists GetMetadataFolder()
    WriteTextFileAtomic gMetaPath, metaJson, "utf-8" ' metadata always UTF-8

    If showConfirmation Then
        MsgBox "Exported: " & gCsvPathAbs, vbInformation
    End If
End Sub


Public Sub SelectCsvAndReload()
    ' Optional utility if you want manual selection without env vars
    Dim p As String
    p = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select CSV")
    If VarType(p) = vbBoolean Then Exit Sub

    gCsvPathOriginalArg = CStr(p)
    gCsvPathAbs = ResolveCsvPath(gCsvPathOriginalArg)
    If Len(gCsvPathAbs) = 0 Or Dir$(gCsvPathAbs) = "" Then
        MsgBox "CSV not found: " & gCsvPathAbs, vbCritical
        Exit Sub
    End If

    gMetaPath = BuildMetaPath(gCsvPathAbs)
    gDelim = DetectDelimiterBestEffort(gCsvPathAbs, ",")
    gEncoding = DetectEncoding(gCsvPathAbs)

    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets(SHEET_DATA)

    LoadCsvIntoSheet gCsvPathAbs, gDelim, gEncoding, wsData
    ApplyDefaultView wsData
End Sub


' =====================
' Path resolution
' =====================

Private Function ResolveCsvPath(ByVal csvArg As String) As String
    ' If absolute -> return canonical absolute
    ' If relative -> base is EXCEL_CSV_CWD; if missing, fallback to workbook folder
    Dim p As String
    p = Trim$(csvArg)

    If Len(p) = 0 Then
        ResolveCsvPath = ""
        Exit Function
    End If

    ' If quoted
    If Left$(p, 1) = """" And Right$(p, 1) = """" Then
        p = Mid$(p, 2, Len(p) - 2)
    End If

    Dim absPath As String
    If IsAbsolutePath(p) Then
        absPath = p
    Else
        Dim baseDir As String
        baseDir = GetEnv("EXCEL_CSV_CWD")
        If Len(baseDir) = 0 Then baseDir = ThisWorkbook.path

        absPath = JoinPath(baseDir, p)
    End If

    ResolveCsvPath = CanonicalizePath(absPath)
End Function

Private Function IsAbsolutePath(ByVal p As String) As Boolean
    p = Trim$(p)

    ' Drive-absolute: C:\...
    If Len(p) >= 3 Then
        If Mid$(p, 2, 1) = ":" And (Mid$(p, 3, 1) = "\" Or Mid$(p, 3, 1) = "/") Then
            IsAbsolutePath = True
            Exit Function
        End If
    End If

    ' UNC: \\server\share\...
    If Len(p) >= 2 Then
        If Left$(p, 2) = "\\" Then
            IsAbsolutePath = True
            Exit Function
        End If
    End If

    IsAbsolutePath = False
End Function


Private Function JoinPath(ByVal dirPath As String, ByVal rel As String) As String
    If Right$(dirPath, 1) = "\" Or Right$(dirPath, 1) = "/" Then
        JoinPath = dirPath & rel
    Else
        JoinPath = dirPath & "\" & rel
    End If
End Function

Private Function CanonicalizePath(ByVal p As String) As String
    On Error GoTo fail
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' GetAbsolutePathName also normalizes ".." and "."
    CanonicalizePath = fso.GetAbsolutePathName(p)
    Exit Function
fail:
    CanonicalizePath = p
End Function


' =====================
' Metadata pathing
' =====================

Private Function GetMetadataFolder() As String
    GetMetadataFolder = JoinPath(ThisWorkbook.path, METADATA_DIR)
End Function

Private Function BuildMetaPath(ByVal csvAbsPath As String) As String
    Dim folder As String
    folder = GetMetadataFolder()

    Dim baseName As String
    baseName = GetFileName(csvAbsPath) ' includes extension

    Dim safeName As String
    safeName = MakeFileNameSafe(baseName)

    Dim h As String
    h = Hex8(Crc32String(csvAbsPath))

    BuildMetaPath = JoinPath(folder, h & "_" & safeName & ".json")
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
' Controls / defaults
' =====================

Private Sub ApplyDefaultView(ByVal ws As Worksheet)
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ws.ListObjects(TABLE_NAME)
    On Error GoTo 0

    If Not lo Is Nothing Then
        lo.TableStyle = "TableStyleMedium2"
    End If

    ws.rows(1).Font.Bold = True
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub


' =====================
' CSV import (text-by-default)
' =====================

Private Sub LoadCsvIntoSheet(ByVal path As String, ByVal delim As String, ByVal enc As String, ByVal ws As Worksheet)
    Dim text As String
    text = ReadTextFileBestEffort(path, enc)

    Do While Len(text) > 0
        If Len(text) >= 2 And Right$(text, 2) = vbCrLf Then
            text = Left$(text, Len(text) - 2)
        ElseIf Right$(text, 1) = vbLf Or Right$(text, 1) = vbCr Then
            text = Left$(text, Len(text) - 1)
        Else
            Exit Do
        End If
    Loop
    
    Dim rows As Variant
    rows = ParseCsvToJagged(text, delim)

    ws.Cells.Clear

    If IsEmpty(rows) Then Exit Sub

    Dim r As Long, c As Long, maxC As Long
    maxC = 0
    For r = LBound(rows) To UBound(rows)
        If Not IsEmpty(rows(r)) Then
            If UBound(rows(r)) > maxC Then maxC = UBound(rows(r))
        End If
    Next r

    Dim nRows As Long, nCols As Long
    nRows = (UBound(rows) - LBound(rows) + 1)
    nCols = maxC + 1

    If nRows = 0 Or nCols = 0 Then Exit Sub

    Dim out() As Variant
    ReDim out(1 To nRows, 1 To nCols)

    For r = 1 To nRows
        Dim rowArr As Variant
        rowArr = rows(LBound(rows) + (r - 1))
        If Not IsEmpty(rowArr) Then
            For c = 1 To nCols
                If (c - 1) <= UBound(rowArr) Then
                    out(r, c) = rowArr(c - 1)
                Else
                    out(r, c) = ""
                End If
            Next c
        End If
    Next r

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, nCols))
    
    ' Default format before writing
    rng.NumberFormat = "General"
    
    rng.Value2 = out
    
    EnsureTable ws, nRows, nCols
    
    PostProcessColumnTypes ws, nRows, nCols
    
    ws.rows(1).Font.Bold = True

End Sub


Private Sub EnsureTable(ByVal ws As Worksheet, ByVal nRows As Long, ByVal nCols As Long)
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ws.ListObjects(TABLE_NAME)
    On Error GoTo 0

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, nCols))

    If lo Is Nothing Then
        Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        lo.name = TABLE_NAME
    Else
        lo.Resize rng
    End If
End Sub


Private Sub PostProcessColumnTypes(ByVal ws As Worksheet, ByVal nRows As Long, ByVal nCols As Long)
    Dim c As Long
    For c = 1 To nCols
        Dim colRng As Range
        Set colRng = ws.Range(ws.Cells(2, c), ws.Cells(nRows, c)) ' skip header

        If ColumnLooksLikeId(colRng) Then
            ws.Columns(c).NumberFormat = "@"
            On Error Resume Next
            colRng.Errors(xlNumberAsText).Ignore = True
            On Error GoTo 0
        Else
            ' Convert numeric-looking strings to real numbers
            ConvertColumnToNumbers colRng
            ws.Columns(c).NumberFormat = "General"
        End If
    Next c
End Sub


Private Function ColumnLooksLikeId(ByVal colRng As Range) As Boolean
    ' Heuristic: if we see leading zeros or very long digit-only values, treat as ID.
    Dim cell As Range
    Dim seen As Long
    For Each cell In colRng.Cells
        Dim s As String
        s = CStr(cell.Value2)
        If Len(s) > 0 Then
            seen = seen + 1
            If Len(s) >= 2 And Left$(s, 1) = "0" And IsAllDigits(s) Then
                ColumnLooksLikeId = True
                Exit Function
            End If
            If Len(s) >= 12 And IsAllDigits(s) Then ' long numbers often should remain text
                ColumnLooksLikeId = True
                Exit Function
            End If
            If seen >= 50 Then Exit For ' sample only
        End If
    Next cell
    ColumnLooksLikeId = False
End Function


Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then
            IsAllDigits = False
            Exit Function
        End If
    Next i
    IsAllDigits = (Len(s) > 0)
End Function


Private Sub ConvertColumnToNumbers(ByVal colRng As Range)
    ' Convert cells that look numeric (no thousands separators; dot decimal) into numbers.
    Dim cell As Range
    For Each cell In colRng.Cells
        Dim s As String
        s = Trim$(CStr(cell.Value2))
        If Len(s) = 0 Then GoTo nextCell

        ' Accept simple numeric patterns; keep others as text
        If LooksNumericSimple(s) Then
            cell.Value2 = CDbl(Replace$(s, ",", ".")) ' tolerate comma decimals
        End If

nextCell:
    Next cell
End Sub


Private Function LooksNumericSimple(ByVal s As String) As Boolean
    ' Very conservative: optional leading '-', digits, optional one decimal separator.
    Dim i As Long, dotSeen As Boolean
    For i = 1 To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)

        If i = 1 And ch = "-" Then
            ' ok
        ElseIf ch >= "0" And ch <= "9" Then
            ' ok
        ElseIf (ch = "." Or ch = ",") And Not dotSeen Then
            dotSeen = True
        Else
            LooksNumericSimple = False
            Exit Function
        End If
    Next i
    LooksNumericSimple = True
End Function


' =====================
' CSV export (normalized quoting + chosen encoding)
' =====================

Private Sub ExportSheetToCsv(ByVal path As String, ByVal delim As String, ByVal enc As String, ByVal ws As Worksheet)
    Dim used As Range
    Set used = ws.UsedRange
    If used Is Nothing Then Exit Sub

    Dim data As Variant
    data = used.Value2

    Dim r As Long, c As Long
    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1)
    nCols = UBound(data, 2)

    Dim lines() As String
    ReDim lines(1 To nRows)

    For r = 1 To nRows
        Dim fields() As String
        ReDim fields(1 To nCols)
        For c = 1 To nCols
            fields(c) = CsvEscape(CStrSafe(data(r, c)), delim)
        Next c
        lines(r) = Join(fields, delim)
    Next r

    Dim csvText As String
    csvText = Join(lines, vbCrLf) & vbCrLf

    WriteTextFileAtomic path, csvText, enc
End Sub

Private Function CStrSafe(ByVal v As Variant) As String
    If IsError(v) Then
        CStrSafe = ""
    ElseIf IsEmpty(v) Or VarType(v) = vbNull Then
        CStrSafe = ""
    Else
        CStrSafe = CStr(v)
    End If
End Function

Private Function CsvEscape(ByVal s As String, ByVal delim As String) As String
    Dim needsQuotes As Boolean
    needsQuotes = (InStr(1, s, delim, vbBinaryCompare) > 0) Or _
                  (InStr(1, s, """", vbBinaryCompare) > 0) Or _
                  (InStr(1, s, vbCr, vbBinaryCompare) > 0) Or _
                  (InStr(1, s, vbLf, vbBinaryCompare) > 0)

    If InStr(1, s, """", vbBinaryCompare) > 0 Then
        s = Replace$(s, """", """""")
    End If

    If needsQuotes Then
        CsvEscape = """" & s & """"
    Else
        CsvEscape = s
    End If
End Function


' =====================
' Metadata (JSON) read/apply/write
' =====================

Private Sub ApplyMetaConfig(ByVal json As String, ByVal ws As Worksheet)
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
        ws.ListObjects(TABLE_NAME).TableStyle = styleName
        On Error GoTo 0
    End If

    ' Freeze panes (row/col)
    Dim fr As Long, fc As Long
    fr = ExtractIntValue(json, "freeze_row", 2)
    fc = ExtractIntValue(json, "freeze_col", 1)

    ' Word wrap
    Dim wraps As Variant
    wraps = ExtractBoolArray(json, "wrap_columns")
    If Not IsEmpty(wraps) Then
        Dim k As Long
        For k = LBound(wraps) To UBound(wraps)
            ws.Columns(k + 1).WrapText = CBool(wraps(k))
        Next k
    End If

    ws.Activate
    ws.Cells(fr, fc).Select
    ActiveWindow.FreezePanes = True
End Sub

Private Function BuildMetaJson(ByVal ws As Worksheet) As String
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
        Erase hParts   ' makes it an unallocated dynamic array
    End If

    ' freeze panes: store the top-left visible cell of the scrollable area
    Dim freezeRow As Long, freezeCol As Long
    freezeRow = 2
    freezeCol = 1
    On Error Resume Next
    If Not ActiveWindow Is Nothing Then
        If ActiveWindow.FreezePanes Then
            freezeRow = ActiveWindow.SplitRow + 1
            freezeCol = ActiveWindow.SplitColumn + 1
        End If
    End If
    On Error GoTo 0

    ' table style
    Dim styleName As String
    styleName = ""
    On Error Resume Next
    styleName = ws.ListObjects(TABLE_NAME).TableStyle
    On Error GoTo 0

    ' wrap settings (per column)
    Dim wrapParts() As String
    If nCols > 0 Then
        ReDim wrapParts(0 To nCols - 1)
        For i = 1 To nCols
            If ws.Columns(i).WrapText Then
                wrapParts(i - 1) = "true"
            Else
                wrapParts(i - 1) = "false"
            End If
        Next i
    Else
        ReDim wrapParts(0 To -1)
    End If

    Dim json As String
    json = "{" & vbCrLf & _
           "  ""csv_path"": " & JsonString(gCsvPathAbs) & "," & vbCrLf & _
           "  ""delimiter"": " & JsonString(gDelim) & "," & vbCrLf & _
           "  ""encoding"": " & JsonString(gEncoding) & "," & vbCrLf & _
           "  ""table_style"": " & JsonString(styleName) & "," & vbCrLf & _
           "  ""freeze_row"": " & CStr(freezeRow) & "," & vbCrLf & _
           "  ""freeze_col"": " & CStr(freezeCol) & "," & vbCrLf & _
           "  ""column_widths"": [" & Join(wParts, ",") & "]," & vbCrLf & _
           "  ""hidden_columns"": [" & Join(hParts, ",") & "]," & vbCrLf & _
           "  ""wrap_columns"": [" & Join(wrapParts, ",") & "]" & vbCrLf & _
           "}" & vbCrLf

    BuildMetaJson = json
End Function

Private Function JsonString(ByVal s As String) As String
    s = Replace$(s, "\", "\\")
    s = Replace$(s, """", "\""")
    s = Replace$(s, vbCr, "\r")
    s = Replace$(s, vbLf, "\n")
    JsonString = """" & s & """"
End Function

Private Function InvariantNumber(ByVal v As Double) As String
    ' Ensure decimal point is '.' regardless of locale
    Dim s As String
    s = CStr(v)
    s = Replace$(s, ",", ".")
    InvariantNumber = s
End Function


' =====================
' Minimal JSON extractors (for our own generated JSON)
' =====================

Private Function ExtractStringValue(ByVal json As String, ByVal key As String) As String
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


' =====================
' Delimiter detection
' =====================

Private Function DetectDelimiterBestEffort(ByVal path As String, ByVal defaultDelim As String) As String
    On Error GoTo fallback
    Dim sample As String
    sample = ReadFirstNChars(path, 4096)

    Dim commas As Long, semis As Long
    commas = CountCharOutsideQuotes(sample, ",")
    semis = CountCharOutsideQuotes(sample, ";")

    If semis > commas And semis > 0 Then
        DetectDelimiterBestEffort = ";"
    ElseIf commas > 0 Then
        DetectDelimiterBestEffort = ","
    Else
        DetectDelimiterBestEffort = defaultDelim
    End If
    Exit Function
fallback:
    DetectDelimiterBestEffort = defaultDelim
End Function

Private Function CountCharOutsideQuotes(ByVal s As String, ByVal ch As String) As Long
    Dim i As Long, inQ As Boolean
    For i = 1 To Len(s)
        Dim c As String
        c = Mid$(s, i, 1)
        If c = """" Then
            inQ = Not inQ
        ElseIf Not inQ And c = ch Then
            CountCharOutsideQuotes = CountCharOutsideQuotes + 1
        End If
    Next i
End Function


' =====================
' CSV parsing (quoted delimiters/newlines, doubled quotes)
' Returns Variant array of rows; each row is Variant array of fields (0-based)
' =====================

Private Function ParseCsvToJagged(ByVal text As String, ByVal delim As String) As Variant
    If Len(text) = 0 Then
        ParseCsvToJagged = Empty
        Exit Function
    End If

    Dim rows() As Variant
    ReDim rows(0 To 0)
    Dim rowCount As Long: rowCount = 0

    Dim fields() As String
    ReDim fields(0 To 0)
    Dim fieldCount As Long: fieldCount = 0

    Dim cur As String: cur = ""
    Dim i As Long
    Dim inQ As Boolean: inQ = False

    For i = 1 To Len(text)
        Dim ch As String
        ch = Mid$(text, i, 1)

        If ch = """" Then
            If inQ And i < Len(text) And Mid$(text, i + 1, 1) = """" Then
                cur = cur & """"
                i = i + 1
            Else
                inQ = Not inQ
            End If

        ElseIf Not inQ And ch = delim Then
            AddField fields, fieldCount, cur
            cur = ""

        ElseIf Not inQ And (ch = vbCr Or ch = vbLf) Then
            If ch = vbCr And i < Len(text) And Mid$(text, i + 1, 1) = vbLf Then
                i = i + 1
            End If

            AddField fields, fieldCount, cur
            cur = ""

            AddRow rows, rowCount, fields, fieldCount

            ReDim fields(0 To 0)
            fieldCount = 0

        Else
            cur = cur & ch
        End If
    Next i

    ' Only add a final row if it contains content OR we already have fields
    ' (e.g. last line ends with a delimiter like "a,")
    If (fieldCount > 0) Or (Len(cur) > 0) Then
        AddField fields, fieldCount, cur
        AddRow rows, rowCount, fields, fieldCount
    End If

    ParseCsvToJagged = rows
End Function

Private Sub AddField(ByRef fields() As String, ByRef fieldCount As Long, ByVal value As String)
    If fieldCount = 0 Then
        fields(0) = value
    Else
        ReDim Preserve fields(0 To fieldCount)
        fields(fieldCount) = value
    End If
    fieldCount = fieldCount + 1
End Sub

Private Sub AddRow(ByRef rows() As Variant, ByRef rowCount As Long, ByRef fields() As String, ByVal fieldCount As Long)
    Dim outFields() As Variant
    Dim i As Long
    ReDim outFields(0 To fieldCount - 1)
    For i = 0 To fieldCount - 1
        outFields(i) = fields(i)
    Next i

    If rowCount = 0 Then
        rows(0) = outFields
    Else
        ReDim Preserve rows(0 To rowCount)
        rows(rowCount) = outFields
    End If
    rowCount = rowCount + 1
End Sub


Private Function IsByteArrayAllocated(ByRef b() As Byte) As Boolean
    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(b)
    ub = UBound(b)
    IsByteArrayAllocated = (Err.Number = 0 And ub >= lb)
    Err.Clear
End Function


' =====================
' File IO + encoding detection
' =====================

Private Function DetectEncoding(ByVal path As String) As String
    Dim b() As Byte
    b = ReadAllBytes(path)
    
    If Not IsByteArrayAllocated(b) Then
        DetectEncoding = "utf-8"
        Exit Function
    End If

    If UBound(b) >= 2 Then
        ' UTF-8 BOM
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then
            DetectEncoding = "utf-8"
            Exit Function
        End If
    End If

    If UBound(b) >= 1 Then
        ' UTF-16 LE BOM
        If b(0) = &HFF And b(1) = &HFE Then
            DetectEncoding = "utf-16le"
            Exit Function
        End If
        ' UTF-16 BE BOM
        If b(0) = &HFE And b(1) = &HFF Then
            DetectEncoding = "utf-16be"
            Exit Function
        End If
    End If

    ' Heuristic: valid UTF-8?
    If LooksLikeUtf8(b) Then
        DetectEncoding = "utf-8"
    Else
        DetectEncoding = "windows-1252"
    End If
End Function

Private Function LooksLikeUtf8(ByRef b() As Byte) As Boolean
    On Error GoTo fail
    Dim i As Long
    i = 0
    Do While i <= UBound(b)
        Dim c As Long
        c = b(i)

        If c < &H80 Then
            i = i + 1
        ElseIf (c And &HE0) = &HC0 Then
            If i + 1 > UBound(b) Then GoTo fail
            If (b(i + 1) And &HC0) <> &H80 Then GoTo fail
            i = i + 2
        ElseIf (c And &HF0) = &HE0 Then
            If i + 2 > UBound(b) Then GoTo fail
            If (b(i + 1) And &HC0) <> &H80 Then GoTo fail
            If (b(i + 2) And &HC0) <> &H80 Then GoTo fail
            i = i + 3
        ElseIf (c And &HF8) = &HF0 Then
            If i + 3 > UBound(b) Then GoTo fail
            If (b(i + 1) And &HC0) <> &H80 Then GoTo fail
            If (b(i + 2) And &HC0) <> &H80 Then GoTo fail
            If (b(i + 3) And &HC0) <> &H80 Then GoTo fail
            i = i + 4
        Else
            GoTo fail
        End If
    Loop

    LooksLikeUtf8 = True
    Exit Function
fail:
    LooksLikeUtf8 = False
End Function

Private Function ReadTextFileBestEffort(ByVal path As String, ByVal enc As String) As String
    If Dir$(path) = "" Then
        ReadTextFileBestEffort = ""
        Exit Function
    End If

    Dim charset As String
    charset = CharsetForEncoding(enc)

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text
    stm.charset = charset
    stm.Open
    stm.LoadFromFile path
    ReadTextFileBestEffort = stm.ReadText
    stm.Close
End Function

Private Function CharsetForEncoding(ByVal enc As String) As String
    Select Case LCase$(enc)
        Case "utf-8"
            CharsetForEncoding = "utf-8"
        Case "windows-1252"
            CharsetForEncoding = "windows-1252"
        Case "utf-16le"
            CharsetForEncoding = "unicode" ' ADODB.Stream: UTF-16LE
        Case "utf-16be"
            ' ADODB.Stream doesn't have a clean UTF-16BE text charset.
            ' Read as binary + manual conversion would be needed. For now:
            CharsetForEncoding = "unicode"
        Case Else
            CharsetForEncoding = "utf-8"
    End Select
End Function

Private Function ReadFirstNChars(ByVal path As String, ByVal n As Long) As String
    Dim s As String
    s = ReadTextFileBestEffort(path, "utf-8")
    If Len(s) > n Then s = Left$(s, n)
    ReadFirstNChars = s
End Function

Private Sub WriteTextFileAtomic(ByVal path As String, ByVal text As String, ByVal enc As String)
    Dim tmpPath As String
    tmpPath = path & ".tmp"

    Dim charset As String
    charset = CharsetForEncoding(enc)

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text
    stm.charset = charset
    stm.Open
    stm.WriteText text
    stm.SaveToFile tmpPath, 2 ' overwrite
    stm.Close

    ' If writing UTF-8, strip BOM for stable diffs
    If LCase$(enc) = "utf-8" Then
        StripUtf8BomIfPresent tmpPath
    End If

    On Error Resume Next
    Kill path
    On Error GoTo 0
    Name tmpPath As path
End Sub

Private Sub StripUtf8BomIfPresent(ByVal filePath As String)
    Dim bytes() As Byte
    bytes = ReadAllBytes(filePath)

    If (Not Not bytes) = 0 Then Exit Sub
    If UBound(bytes) >= 2 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            Dim outBytes() As Byte
            Dim i As Long
            ReDim outBytes(0 To UBound(bytes) - 3)
            For i = 3 To UBound(bytes)
                outBytes(i - 3) = bytes(i)
            Next i
            WriteAllBytes filePath, outBytes
        End If
    End If
End Sub

Private Function ReadAllBytes(ByVal filePath As String) As Byte()
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' binary
    stm.Open
    stm.LoadFromFile filePath
    ReadAllBytes = stm.Read
    stm.Close
End Function

Private Sub WriteAllBytes(ByVal filePath As String, ByRef bytes() As Byte)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' binary
    stm.Open
    stm.Write bytes
    stm.SaveToFile filePath, 2 ' overwrite
    stm.Close
End Sub

Private Sub EnsureFolderExists(ByVal folderPath As String)
    If Len(folderPath) = 0 Then Exit Sub
    If Dir$(folderPath, vbDirectory) <> "" Then Exit Sub
    On Error GoTo done
    MkDir folderPath
done:
End Sub


' =====================
' Environment helper
' =====================

Private Function GetEnv(ByVal name As String) As String
    GetEnv = ""  ' default if missing or error

    On Error GoTo done
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    GetEnv = CStr(wsh.Environment("PROCESS")(name))
done:
End Function



' =====================
' CRC32 hashing (for metadata filename)
' =====================

Private Function Crc32String(ByVal s As String) As Long
    Dim bytes() As Byte
    bytes = StrToUtf8Bytes(s)
    Crc32String = Crc32Bytes(bytes)
End Function


Private Function StrToUtf8Bytes(ByVal s As String) As Byte()
    ' Write string as UTF-8 to ADODB.Stream then read bytes back as Byte()
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2              ' adTypeText
    stm.charset = "utf-8"
    stm.Open
    stm.WriteText s

    stm.Position = 0
    stm.Type = 1              ' adTypeBinary

    Dim v As Variant
    v = stm.Read              ' Variant(byte array)

    stm.Close

    Dim b() As Byte
    b = v

    ' Strip UTF-8 BOM if present
    If HasUtf8Bom(b) Then
        b = SliceBytes(b, 3)
    End If

    StrToUtf8Bytes = b
End Function


Private Function HasUtf8Bom(ByRef b() As Byte) As Boolean
    On Error GoTo no
    If UBound(b) >= 2 Then
        HasUtf8Bom = (b(0) = &HEF And b(1) = &HBB And b(2) = &HBF)
        Exit Function
    End If
no:
    HasUtf8Bom = False
End Function


Private Function SliceBytes(ByRef b() As Byte, ByVal offset As Long) As Byte()
    ' Returns b[offset..end]
    Dim out() As Byte
    Dim i As Long, n As Long

    n = (UBound(b) - offset + 1)
    If n <= 0 Then
        ReDim out(0 To -1)
        SliceBytes = out
        Exit Function
    End If

    ReDim out(0 To n - 1)
    For i = 0 To n - 1
        out(i) = b(offset + i)
    Next i

    SliceBytes = out
End Function



Private Function Crc32Bytes(ByRef b() As Byte) As Long
    Dim crc As Long
    crc = &HFFFFFFFF

    Dim i As Long, j As Long
    For i = LBound(b) To UBound(b)
        crc = crc Xor b(i)
        For j = 1 To 8
            If (crc And 1) <> 0 Then
                crc = ((crc And &HFFFFFFFE) \ 2) Xor &HEDB88320
            Else
                crc = (crc And &HFFFFFFFE) \ 2
            End If
        Next j
    Next i

    Crc32Bytes = Not crc
End Function

Private Function Hex8(ByVal v As Long) As String
    Dim s As String
    s = Hex$(v)
    Hex8 = Right$("00000000" & s, 8)
End Function






