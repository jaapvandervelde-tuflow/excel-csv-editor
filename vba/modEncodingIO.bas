Attribute VB_Name = "modEncodingIO"
Option Explicit

' =====================
' File IO + encoding detection
' =====================

Public Function DetectEncoding(ByVal path As String) As String
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

Public Function ReadTextFileBestEffort(ByVal path As String, ByVal enc As String) As String
    If Dir$(path) = "" Then
        ReadTextFileBestEffort = ""
        Exit Function
    End If

    Select Case LCase$(enc)
        Case "utf-16be"
            ReadTextFileBestEffort = ReadUtf16BeText(path)
        Case Else
            ReadTextFileBestEffort = ReadTextViaAdodb(path, CharsetForEncoding(enc))
    End Select
End Function

Private Function ReadTextViaAdodb(ByVal path As String, ByVal charset As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text
    stm.charset = charset
    stm.Open
    stm.LoadFromFile path
    ReadTextViaAdodb = stm.ReadText
    stm.Close
End Function

Private Function ReadUtf16BeText(ByVal path As String) As String
    Dim bytes() As Byte
    bytes = ReadAllBytes(path)
    If Not IsByteArrayAllocated(bytes) Then
        ReadUtf16BeText = ""
        Exit Function
    End If

    Dim offset As Long
    offset = 0
    ' If BOM present, skip it
    If UBound(bytes) >= 1 Then
        If bytes(0) = &HFE And bytes(1) = &HFF Then
            offset = 2
        End If
    End If

    Dim n As Long
    n = (UBound(bytes) - offset + 1)
    If n <= 0 Then
        ReadUtf16BeText = ""
        Exit Function
    End If

    ' Swap BE -> LE so VBA can consume as UTF-16LE.
    Dim le() As Byte
    le = SwapUtf16ByteOrder(bytes, offset)

    ' Convert LE bytes to String using ADODB.Stream as "unicode" (UTF-16LE).
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' binary
    stm.Open
    stm.Write le
    stm.Position = 0
    stm.Type = 2 ' text
    stm.charset = "unicode"
    ReadUtf16BeText = stm.ReadText
    stm.Close
End Function

Private Function CharsetForEncoding(ByVal enc As String) As String
    Select Case LCase$(enc)
        Case "utf-8"
            CharsetForEncoding = "utf-8"
        Case "windows-1252"
            CharsetForEncoding = "windows-1252"
        Case "utf-16le"
            ' ADODB.Stream: UTF-16LE
            CharsetForEncoding = "unicode"
        Case "utf-16be"
            ' handled separately (binary swap)
            CharsetForEncoding = "unicode"
        Case Else
            CharsetForEncoding = "utf-8"
    End Select
End Function

Public Sub WriteTextFileAtomic(ByVal path As String, ByVal text As String, ByVal enc As String)
    Dim tmpPath As String
    tmpPath = path & ".tmp"

    Select Case LCase$(enc)
        Case "utf-16be"
            WriteUtf16BeAtomic tmpPath, text
        Case Else
            WriteTextViaAdodb tmpPath, text, CharsetForEncoding(enc)
    End Select

    ' If writing UTF-8, strip BOM for stable diffs
    If LCase$(enc) = "utf-8" Then
        StripUtf8BomIfPresent tmpPath
    End If

    On Error Resume Next
    Kill path
    On Error GoTo 0
    Name tmpPath As path
End Sub

Private Sub WriteTextViaAdodb(ByVal path As String, ByVal text As String, ByVal charset As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text
    stm.charset = charset
    stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2 ' overwrite
    stm.Close
End Sub

Private Sub WriteUtf16BeAtomic(ByVal tmpPath As String, ByVal text As String)
    ' Write as UTF-16LE first, then swap to BE and add BOM (FE FF).
    Dim leBytes() As Byte
    leBytes = StringToUtf16LeBytes(text) ' includes BOM (FF FE)

    ' Remove LE BOM if present before swapping, then prepend BE BOM after swap.
    Dim offset As Long
    offset = 0
    If IsByteArrayAllocated(leBytes) Then
        If UBound(leBytes) >= 1 Then
            If leBytes(0) = &HFF And leBytes(1) = &HFE Then
                offset = 2
            End If
        End If
    End If

    Dim beBody() As Byte
    beBody = SwapUtf16ByteOrder(leBytes, offset)

    Dim out() As Byte
    If Not IsByteArrayAllocated(beBody) Then
        ReDim out(0 To 1)
        out(0) = &HFE: out(1) = &HFF
    Else
        ReDim out(0 To UBound(beBody) + 2)
        out(0) = &HFE: out(1) = &HFF

        Dim i As Long
        For i = 0 To UBound(beBody)
            out(i + 2) = beBody(i)
        Next i
    End If

    WriteAllBytes tmpPath, out
End Sub

Private Function StringToUtf16LeBytes(ByVal s As String) As Byte()
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.charset = "unicode" ' UTF-16LE with BOM
    stm.Open
    stm.WriteText s

    stm.Position = 0
    stm.Type = 1

    Dim v As Variant
    v = stm.Read
    stm.Close

    Dim b() As Byte
    b = v
    StringToUtf16LeBytes = b
End Function

Private Function SwapUtf16ByteOrder(ByRef bytes() As Byte, ByVal offset As Long) As Byte()
    Dim n As Long
    If Not IsByteArrayAllocated(bytes) Then
        ReDim SwapUtf16ByteOrder(0 To -1)
        Exit Function
    End If

    n = UBound(bytes) - offset + 1
    If n <= 0 Then
        ReDim SwapUtf16ByteOrder(0 To -1)
        Exit Function
    End If

    ' Ensure even number of bytes for UTF-16 code units.
    If (n Mod 2) <> 0 Then n = n - 1
    If n <= 0 Then
        ReDim SwapUtf16ByteOrder(0 To -1)
        Exit Function
    End If

    Dim out() As Byte
    ReDim out(0 To n - 1)

    Dim i As Long
    For i = 0 To n - 1 Step 2
        out(i) = bytes(offset + i + 1)
        out(i + 1) = bytes(offset + i)
    Next i

    SwapUtf16ByteOrder = out
End Function

Private Sub StripUtf8BomIfPresent(ByVal filePath As String)
    Dim bytes() As Byte
    bytes = ReadAllBytes(filePath)

    If Not IsByteArrayAllocated(bytes) Then Exit Sub
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

' =====================
' Binary IO helpers
' =====================

Public Function ReadAllBytes(ByVal filePath As String) As Byte()
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' binary
    stm.Open
    stm.LoadFromFile filePath
    ReadAllBytes = stm.Read
    stm.Close
End Function

Public Sub WriteAllBytes(ByVal filePath As String, ByRef bytes() As Byte)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' binary
    stm.Open
    stm.Write bytes
    stm.SaveToFile filePath, 2 ' overwrite
    stm.Close
End Sub

Public Function IsByteArrayAllocated(ByRef b() As Byte) As Boolean
    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(b)
    ub = UBound(b)
    IsByteArrayAllocated = (Err.Number = 0 And ub >= lb)
    Err.Clear
End Function

Public Sub EnsureFolderExists(ByVal folderPath As String)
    If Len(folderPath) = 0 Then Exit Sub
    If Dir$(folderPath, vbDirectory) <> "" Then Exit Sub
    On Error GoTo done
    MkDir folderPath
done:
End Sub

