Attribute VB_Name = "ShapeProperties"

Dim refTop, refLeft, refHeight, refWidth, refRotation As Single
Dim refMarginTop, refMarginBottom, refMarginLeft, refMarginRight As Single
Dim refAdjustments(7) As Single

Sub ButtonLearnMore(control As IRibbonControl)
    ActivePresentation.FollowHyperlink ("https://github.com/leliukhin/ppt-macros")
End Sub

Sub ButtonCopyProperties(control As IRibbonControl)
    
    Dim referenceShape As Shape
    Dim adjCount As Integer
    On Error Resume Next
    Err.Clear

    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)

    If Err <> 0 Then
        MsgBox "Select a shape."
    Else
        refTop = referenceShape.Top
        refLeft = referenceShape.Left
        refHeight = referenceShape.Height
        refWidth = referenceShape.Width
        refRotation = referenceShape.Rotation
        refMarginTop = referenceShape.TextFrame.marginTop
        refMarginBottom = referenceShape.TextFrame.marginBottom
        refMarginLeft = referenceShape.TextFrame.marginLeft
        refMarginRight = referenceShape.TextFrame.marginRight
        
        adjCount = referenceShape.Adjustments.Count
        
        For i = 1 To adjCount
            refAdjustments(i - 1) = referenceShape.Adjustments.Item(i)
        Next
        
    End If
End Sub

Sub ButtonPasteTopLeft(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.Top = refTop
            shp.Left = refLeft
        End If
    Next shp
    
End Sub

Sub ButtonPasteSize(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.Height = refHeight
            shp.Width = refWidth
        End If
    Next shp
    
End Sub

Sub ButtonPasteTop(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.Top = refTop
        End If
    Next shp
    
End Sub

Sub ButtonPasteLeft(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.Left = refLeft
        End If
    Next shp
    
End Sub

Sub ButtonPasteBottom(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.Top = refTop + refHeight - shp.Height
        End If
    Next shp
    
End Sub

Sub ButtonPasteRight(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.Left = refLeft + refWidth - shp.Width
        End If
    Next shp
    
End Sub

Sub ButtonPasteCenter(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.Left = refLeft + refWidth / 2 - shp.Width / 2
        End If
    Next shp
    
End Sub

Sub ButtonPasteMiddle(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.Top = refTop + refHeight / 2 - shp.Height / 2
        End If
    Next shp
    
End Sub

Sub ButtonPasteCenterMiddle(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.Top = refTop + refHeight / 2 - shp.Height / 2
            shp.Left = refLeft + refWidth / 2 - shp.Width / 2
        End If
    Next shp
    
End Sub

Sub ButtonPasteRotation(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            If shp.Type = msoLine Then
                ' If copied line is straight
                If refHeight = 0 Or refWidth = 0 Then
                    shp.Height = refHeight
                    shp.Width = refWidth
                Else
                    Dim lineLength As Single
                    lineLength = Sqr(shp.Width ^ 2 + shp.Height ^ 2)
                    
                    shp.Height = Sqr((lineLength ^ 2) / ((refWidth / refHeight) ^ 2 + 1))
                    shp.Width = refWidth / refHeight * shp.Height
                End If
            End If
            shp.Rotation = refRotation
        End If
    Next shp
    
End Sub

Sub ButtonPasteHeight(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.Height = refHeight
        End If
    Next shp
    
End Sub

Sub ButtonPasteWidth(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.Width = refWidth
        End If
    Next shp
    
End Sub

Sub ButtonPasteArea(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    targetLogoArea = refHeight * refWidth
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
        shp.LockAspectRatio = msoTrue
        originalLogoArea = shp.Height * shp.Width
            
        shp.Width = shp.Width * Sqr(targetLogoArea / originalLogoArea)
        End If
    Next shp
    
End Sub

Sub ButtonPasteAdjustments(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear

    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            For i = 1 To shp.Adjustments.Count
                shp.Adjustments.Item(i) = refAdjustments(i - 1)
            Next
        End If
    Next shp
    
End Sub

Sub ButtonPasteAll(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear

    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.Height = refHeight
            shp.Width = refWidth
            shp.Top = refTop
            shp.Left = refLeft
        End If
    Next shp
    
End Sub

Sub ButtonPasteMargins(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear

    If refHeight = 0 And refWidth = 0 Then
        MsgBox "Copy shape properties first."
        Exit Sub
    End If
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.TextFrame.marginTop = refMarginTop
            shp.TextFrame.marginBottom = refMarginBottom
            shp.TextFrame.marginLeft = refMarginLeft
            shp.TextFrame.marginRight = refMarginRight
        End If
    Next shp
    
End Sub

Sub ButtonRotateIncrClockwise(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.Rotation = shp.Rotation + 0.3
        End If
    Next shp
    
End Sub

Sub ButtonRotateIncrCounterClockwise(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            shp.Rotation = shp.Rotation - 0.3
        End If
    Next shp
    
End Sub

Sub ButtonRotateStraighten(control As IRibbonControl)

    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            If shp.Type = msoLine Then
                If shp.Height > shp.Width Then
                    shp.Width = 0
                Else
                    shp.Height = 0
                End If
            End If
            shp.Rotation = 0
        End If
    Next shp
    
End Sub