Attribute VB_Name = "modControls"
Public Sub UpdateCsvPathLabel()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Controls")

    Dim p As String
    p = modCsvEditor.CurrentCsvPathAbs()

    Dim txt As String
    If Len(p) = 0 Then
        txt = "(not loaded)"
    Else
        txt = p
    End If

    Dim shp As Shape
    Set shp = ws.Shapes("lblCsvPath")   ' ensure you renamed it to this

    ' Forms label / many shapes: use TextFrame
    shp.TextFrame.Characters.text = txt
End Sub



Private Sub AutoFitAndClampLabel(ByVal shp As Shape)
    ' Forms label as Shape: use TextFrame2 autosize, but clamp to a max width
    Dim maxWidth As Double
    maxWidth = 700 ' adjust to your layout

    With shp.TextFrame2
        .AutoSize = msoAutoSizeShapeToFitText
        .WordWrap = msoFalse
    End With

    If shp.Width > maxWidth Then
        shp.Width = maxWidth
        shp.TextFrame2.AutoSize = msoAutoSizeNone
        shp.TextFrame2.WordWrap = msoTrue
    End If
End Sub
