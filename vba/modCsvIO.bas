Attribute VB_Name = "modCsvIO"
Option Explicit

' =====================
' Controls / defaults
' =====================

Public Sub ApplyDefaultView(ByVal ws As Worksheet, ByVal tableName As String)
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ws.ListObjects(tableName)
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

Public Sub LoadCsvIntoSheet(ByVal path As String, ByVal delim As String, ByVal enc As String, ByVal ws As Worksheet, ByVal tableName As String)
    Dim text As String
    text = modEncodingIO.ReadTextFileBestEffort(path, enc)

    ' Trim trailing line endings
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
    rng.NumberFormat = "General"
    rng.Value2 = out

    EnsureTable ws, nRows, nCols, tableName
    PostProcessColumnTypes ws, nRows, nCols

    ws.rows(1).Font.Bold = True
End Sub

Private Sub EnsureTable(ByVal ws As Worksheet, ByVal nRows As Long, ByVal nCols As Long, ByVal tableName As String)
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, nCols))

    If lo Is Nothing Then
        Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        lo.name = tableName
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
            ConvertColumnToNumbers colRng
            ws.Columns(c).NumberFormat = "General"
        End If
    Next c
End Sub

Private Function ColumnLooksLikeId(ByVal colRng As Range) As Boolean
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
            If Len(s) >= 12 And IsAllDigits(s) Then
                ColumnLooksLikeId = True
                Exit Function
            End If
            If seen >= 50 Then Exit For
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
    Dim cell As Range
    For Each cell In colRng.Cells
        Dim s As String
        s = Trim$(CStr(cell.Value2))
        If Len(s) = 0 Then GoTo nextCell

        If LooksNumericSimple(s) Then
            cell.Value2 = CDbl(Replace$(s, ",", "."))
        End If
nextCell:
    Next cell
End Sub

Private Function LooksNumericSimple(ByVal s As String) As Boolean
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

Public Sub ExportSheetToCsv(ByVal path As String, ByVal delim As String, ByVal enc As String, ByVal ws As Worksheet, ByVal tableName As String)
    Dim src As Range

    ' prefer the table range if present.
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    If Not lo Is Nothing Then
        Set src = lo.Range
    Else
        Set src = ws.UsedRange
    End If

    If src Is Nothing Then Exit Sub

    Dim data As Variant
    data = src.Value2

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

    modEncodingIO.WriteTextFileAtomic path, csvText, enc
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
' Delimiter detection (byte-based sample)
' =====================

Public Function DetectDelimiterBestEffortBytes(ByVal path As String, ByVal defaultDelim As String) As String
    On Error GoTo fallback

    Dim b() As Byte
    b = modEncodingIO.ReadAllBytes(path)
    If Not modEncodingIO.IsByteArrayAllocated(b) Then
        DetectDelimiterBestEffortBytes = defaultDelim
        Exit Function
    End If

    Dim maxN As Long
    maxN = 4096
    If UBound(b) + 1 < maxN Then maxN = UBound(b) + 1

    Dim commas As Long, semis As Long
    commas = CountByteCharOutsideQuotes(b, maxN, Asc(","))
    semis = CountByteCharOutsideQuotes(b, maxN, Asc(";"))

    If semis > commas And semis > 0 Then
        DetectDelimiterBestEffortBytes = ";"
    ElseIf commas > 0 Then
        DetectDelimiterBestEffortBytes = ","
    Else
        DetectDelimiterBestEffortBytes = defaultDelim
    End If
    Exit Function

fallback:
    DetectDelimiterBestEffortBytes = defaultDelim
End Function

Private Function CountByteCharOutsideQuotes(ByRef b() As Byte, ByVal n As Long, ByVal target As Byte) As Long
    Dim i As Long, inQ As Boolean
    Dim quoteB As Byte
    quoteB = Asc("""")

    For i = 0 To n - 1
        Dim c As Byte
        c = b(i)
        If c = quoteB Then
            inQ = Not inQ
        ElseIf Not inQ And c = target Then
            CountByteCharOutsideQuotes = CountByteCharOutsideQuotes + 1
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


