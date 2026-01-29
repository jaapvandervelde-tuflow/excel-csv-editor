Attribute VB_Name = "modHash"
Option Explicit

' =====================
' CRC32 hashing
' =====================

Public Function Crc32String(ByVal s As String) As Long
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

Public Function Hex8(ByVal v As Long) As String
    Dim s As String
    s = Hex$(v)
    Hex8 = Right$("00000000" & s, 8)
End Function

