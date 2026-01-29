VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private mDevMode As Boolean

Public Sub SetDevMode(ByVal enabled As Boolean)
    mDevMode = enabled
End Sub

Public Property Get DevModeEnabled() As Boolean
    DevModeEnabled = mDevMode Or (LCase$(GetEnv("CSV_EDITOR_DEV")) = "1")
End Property


Private Sub Workbook_Open()
    On Error Resume Next
    modCsvSession.ReloadFromEnv
    On Error GoTo 0

    ' After reload, treat workbook as clean
    Me.Saved = True
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Not DevModeEnabled() Then
        Cancel = True

        On Error GoTo exportFail
        modCsvSession.ExportToOriginalCsv

        ' Mark clean after successful export
        Me.Saved = True
        Exit Sub

exportFail:
        ' Keep dirty so user still gets prompted on close
        Me.Saved = False
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If DevModeEnabled() Then Exit Sub

    ' If clean, no prompt; if dirty, 3-way dialog
    If Me.Saved = False Then
        Dim res As VbMsgBoxResult
        res = MsgBox( _
            "You have unsaved changes, save before closing?" & vbCrLf & vbCrLf & _
            "Yes = Save and Close" & vbCrLf & _
            "No = Close and Lose Changes" & vbCrLf & _
            "Cancel = Do not Close", _
            vbExclamation Or vbYesNoCancel, "CSV Editor")

        Select Case res
            Case vbYes
                ' Close, save
                On Error GoTo saveFail
                modCsvSession.ExportToOriginalCsv False
                Me.Saved = True

            Case vbNo
                ' Close, discard
                Me.Saved = True

            Case vbCancel
                ' Do not close
                Cancel = True
        End Select
    End If
    Exit Sub

saveFail:
    Cancel = True
    MsgBox "Export failed; close cancelled.", vbCritical, "CSV Editor"
End Sub


Private Function GetEnv(ByVal name As String) As String
    On Error GoTo done
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    GetEnv = CStr(wsh.Environment("PROCESS")(name))
done:
End Function

