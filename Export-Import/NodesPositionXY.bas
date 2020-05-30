Attribute VB_Name = "NodesPositionXY"
Sub CurveNodesPosition()
Dim s As Shape
Dim n As Integer

If ActiveShape Is Nothing Then
    MsgBox "Nothing selected!", vbCritical
    Exit Sub
End If

For Each s In ActiveSelection.Shapes
    If s.Type = cdrCurveShape Then
        If s.Curve.Selection.Count <> o Then n = n + 1
    Else
        MsgBox "Only curves must be selected!", vbCritical
        Exit Sub
    End If
Next s


If n > o Then
    NodesPos.Show
Else
    MsgBox "Select any node!", vbCritical
    Exit Sub
End If

End Sub
