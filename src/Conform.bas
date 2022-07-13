Attribute VB_Name = "Conform"

Sub ButtonConformTop(control As IRibbonControl)

    Dim referenceShape, shp As Shape
    Dim refTop As Single
    On Error Resume Next
    Err.Clear

    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)
    refTop = referenceShape.Top
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select two or more shapes."
        Else
            shp.Top = refTop
        End If
    Next shp
    
End Sub

Sub ButtonConformBottom(control As IRibbonControl)

    Dim referenceShape, shp As Shape
    Dim refTop, refHeight As Single
    On Error Resume Next
    Err.Clear

    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)
    refTop = referenceShape.Top
    refHeight = referenceShape.Height
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select two or more shapes."
        Else
            shp.Top = refTop + refHeight - shp.Height
        End If
    Next shp
    
End Sub

Sub ButtonConformLeft(control As IRibbonControl)

    Dim referenceShape, shp As Shape
    Dim refLeft As Single
    On Error Resume Next
    Err.Clear

    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)
    refLeft = referenceShape.Left
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select two or more shapes."
        Else
            shp.Left = refLeft
        End If
    Next shp
    
End Sub

Sub ButtonConformRight(control As IRibbonControl)

    Dim referenceShape, shp As Shape
    Dim refLeft, refWidth As Single
    On Error Resume Next
    Err.Clear

    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)
    refLeft = referenceShape.Left
    refWidth = referenceShape.Width
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select two or more shapes."
        Else
            shp.Left = refLeft + refWidth - shp.Width
        End If
    Next shp
    
End Sub

Sub ButtonConformCenter(control As IRibbonControl)

    Dim referenceShape, shp As Shape
    Dim refLeft, refWidth As Single
    On Error Resume Next
    Err.Clear
    
    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)
    refLeft = referenceShape.Left
    refWidth = referenceShape.Width
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select two or more shapes."
        Else
            shp.Left = refLeft + refWidth / 2 - shp.Width / 2
        End If
    Next shp
    
End Sub

Sub ButtonConformMiddle(control As IRibbonControl)

    Dim referenceShape, shp As Shape
    Dim refTop, refHeight As Single
    On Error Resume Next
    Err.Clear
    
    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)
    refTop = referenceShape.Top
    refHeight = referenceShape.Height
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select two or more shapes."
        Else
            shp.Top = refTop + refHeight / 2 - shp.Height / 2
        End If
    Next shp
    
End Sub

Sub ButtonConformMidpoint(control As IRibbonControl)

    Dim referenceShape, shp As Shape
    Dim refTop, refHeight, refLeft, refWidth As Single
    On Error Resume Next
    Err.Clear
    
    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)
    refTop = referenceShape.Top
    refHeight = referenceShape.Height
    refLeft = referenceShape.Left
    refWidth = referenceShape.Width
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select two or more shapes."
        Else
            shp.Top = refTop + refHeight / 2 - shp.Height / 2
            shp.Left = refLeft + refWidth / 2 - shp.Width / 2
        End If
    Next shp
    
End Sub

Sub ButtonConformHeight(control As IRibbonControl)

    Dim referenceShape, shp As Shape
    Dim refHeight As Single
    On Error Resume Next
    Err.Clear

    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)
    refHeight = referenceShape.Height
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select two or more shapes."
        Else
            shp.Height = refHeight
        End If
    Next shp
    
End Sub

Sub ButtonConformWidth(control As IRibbonControl)

    Dim referenceShape, shp As Shape
    Dim refWidth As Single
    On Error Resume Next
    Err.Clear

    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)
    refWidth = referenceShape.Width
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select two or more shapes."
        Else
            shp.Width = refWidth
        End If
    Next shp
    
End Sub

Sub ButtonConformSize(control As IRibbonControl)

    Dim referenceShape, shp As Shape
    Dim refHeight, refWidth As Single
    On Error Resume Next
    Err.Clear

    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)
    refHeight = referenceShape.Height
    refWidth = referenceShape.Width
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select two or more shapes."
        Else
            shp.Height = refHeight
            shp.Width = refWidth
        End If
    Next shp
    
End Sub

Sub ButtonConformRotation(control As IRibbonControl)

    Dim referenceShape, shp As Shape
    Dim refRotation, refHeight, refWidth As Single
    On Error Resume Next
    Err.Clear

    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)
    refRotation = referenceShape.Rotation
    refHeight = referenceShape.Height
    refWidth = referenceShape.Width
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select two or more shapes."
        Else
            If shp.Type = msoLine Then
                ' If reference line is straight
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

Sub ButtonConformAdjustments(control As IRibbonControl)

    Dim referenceShape, shp As Shape
    Dim numberAdjustments As Integer
    Dim conform_Adj(7) As Single
    On Error Resume Next
    Err.Clear

    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)

    ' Copy adjustment values
    numberAdjustments = referenceShape.Adjustments.Count
    conform_Adj(0) = 12345678

    For i = 1 To numberAdjustments
        conform_Adj(i - 1) = referenceShape.Adjustments.Item(i)
    Next

    If conform_Adj(0) = 12345678 Then
        Exit Sub
    End If
    
    ' Paste adjustment values
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select two or more shapes."
        Else
            For i = 1 to numberAdjustments
                shp.Adjustments.Item(i) = conform_Adj(i - 1)
            Next
        End If
    Next shp
    
End Sub

Sub ButtonConformArea(control As IRibbonControl)

    Dim referenceShape, shp, testshp As Shape
    Dim refArea, aspectRatio, originalLogoArea As Single
    On Error Resume Next
    Err.Clear
    
    ' Check if logos are selected
    Set testshp = ActiveWindow.Selection.ShapeRange(2)
    If Err <> 0 Then
        MsgBox "Select two or more images."
        Exit Sub
    End If

    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)
    refArea = referenceShape.Width * referenceShape.Height
        
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.LockAspectRatio = msoTrue
        aspectRatio = shp.Height / shp.Width
        originalLogoArea = shp.Height * shp.Width
            
        shp.Width = shp.Width * Sqr(refArea / originalLogoArea)
    Next shp

End Sub
