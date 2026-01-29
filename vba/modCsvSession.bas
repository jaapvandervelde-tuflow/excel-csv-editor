Attribute VB_Name = "modCsvSession"
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

' Sheet/table names
Private Const SHEET_DATA As String = "Data"
Private Const TABLE_NAME As String = "Table1"

Public Function CurrentCsvPathAbs() As String
    CurrentCsvPathAbs = gCsvPathAbs
End Function

Public Function CurrentDelimiter() As String
    CurrentDelimiter = gDelim
End Function

Public Function CurrentEncoding() As String
    CurrentEncoding = gEncoding
End Function

' =====================
' Public entry points
' =====================

Public Sub ReloadFromEnv()
    Dim pArg As String
    pArg = modCsvSession_GetEnv("EXCEL_CSV_PATH")
    If Len(pArg) = 0 Then
        If Application.UserControl And Not ThisWorkbook.DevModeEnabled Then
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

    gMetaPath = modMeta.BuildMetaPath(gCsvPathAbs)

    ' Load meta first (may set delimiter/encoding/style settings)
    Dim metaJson As String
    metaJson = modEncodingIO.ReadTextFileBestEffort(gMetaPath, "utf-8")

    Dim metaExists As Boolean
    metaExists = (Len(metaJson) > 0)

    ' Delimiter selection:
    ' 1) EXCEL_CSV_DELIM (if set)
    ' 2) metadata (if present)
    ' 3) auto-detect (comma vs semicolon)
    ' 4) default comma
    Dim envDelim As String
    envDelim = modCsvSession_GetEnv("EXCEL_CSV_DELIM")
    If Len(envDelim) > 0 Then
        gDelim = Left$(envDelim, 1)
    ElseIf metaExists Then
        Dim md As String
        md = modMeta.ExtractStringValue(metaJson, "delimiter")
        If Len(md) > 0 Then gDelim = Left$(md, 1)
    End If
    If Len(gDelim) = 0 Then
        gDelim = modCsvIO.DetectDelimiterBestEffortBytes(gCsvPathAbs, ",")
    End If

    ' Encoding selection:
    ' 1) metadata (if present) to preserve previous detected encoding
    ' 2) detect from file
    gEncoding = ""
    If metaExists Then
        gEncoding = modMeta.ExtractStringValue(metaJson, "encoding")
    End If
    If Len(gEncoding) = 0 Then
        gEncoding = modEncodingIO.DetectEncoding(gCsvPathAbs)
    End If

    ' Load CSV to Data sheet
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets(SHEET_DATA)

    modCsvIO.LoadCsvIntoSheet gCsvPathAbs, gDelim, gEncoding, wsData, TABLE_NAME

    ' Apply meta formatting/settings
    If metaExists Then
        modMeta.ApplyMetaConfig metaJson, wsData, TABLE_NAME
    Else
        modCsvIO.ApplyDefaultView wsData, TABLE_NAME
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

    ' Export from the table range if present, else UsedRange.
    modCsvIO.ExportSheetToCsv gCsvPathAbs, gDelim, gEncoding, wsData, TABLE_NAME

    ' Always write metadata after export (captures current view settings)
    Dim metaJson As String
    metaJson = modMeta.BuildMetaJson(wsData, TABLE_NAME, gCsvPathAbs, gDelim, gEncoding)

    modMeta.EnsureMetadataFolder
    modEncodingIO.WriteTextFileAtomic gMetaPath, metaJson, "utf-8" ' metadata always UTF-8

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

    gMetaPath = modMeta.BuildMetaPath(gCsvPathAbs)
    gDelim = modCsvIO.DetectDelimiterBestEffortBytes(gCsvPathAbs, ",")
    gEncoding = modEncodingIO.DetectEncoding(gCsvPathAbs)

    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets(SHEET_DATA)

    modCsvIO.LoadCsvIntoSheet gCsvPathAbs, gDelim, gEncoding, wsData, TABLE_NAME
    modCsvIO.ApplyDefaultView wsData, TABLE_NAME
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
        baseDir = modCsvSession_GetEnv("EXCEL_CSV_CWD")
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
    CanonicalizePath = fso.GetAbsolutePathName(p)
    Exit Function
fail:
    CanonicalizePath = p
End Function

' =====================
' Environment helper
' =====================

Private Function modCsvSession_GetEnv(ByVal name As String) As String
    modCsvSession_GetEnv = ""  ' default if missing or error
    On Error GoTo done
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    modCsvSession_GetEnv = CStr(wsh.Environment("PROCESS")(name))
done:
End Function


