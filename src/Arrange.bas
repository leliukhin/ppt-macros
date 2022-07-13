Attribute VB_Name = "Arrange"

Sub ButtonArrangeVertical(control As IRibbonControl)

    Dim shp As Shape
    Dim topWatermark, increment, topValue, heightValue, rotationAdj As Single
    On Error Resume Next
    Err.Clear

    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Select two or more shapes."
        Exit Sub
    End If

    topWatermark = 12345678
    increment = InputBox("Enter vertical gap in points." & vbNewLine & "(72 points = 1 inch; negative values apply)")
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Rotation = 270 Or shp.Rotation = 90 Then
            ' Determine top value for rotated shapes
            If shp.Height <> shp.Width Then
                topValue = shp.Top + (shp.Height - shp.Width) / 2
            Else
                topValue = shp.Top
            End If
        Else
            topValue = shp.Top
        End If
        
        If topValue < topWatermark Then
            topWatermark = topValue
        End If
    Next shp
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Enter a numeric gap."
            Exit Sub
        Else
            ' Determine effective height of shape for calculations
            If shp.Rotation = 270 Or shp.Rotation = 90 Then
                heightValue = shp.Width
                rotationAdj = (shp.Width - shp.Height) / 2
            Else
                heightValue = shp.Height
                rotationAdj = 0
            End If
                
            ' Arrange shapes
            shp.Top = topWatermark + rotationAdj
            topWatermark = topWatermark + heightValue + increment
        End If
    Next shp
    
End Sub

Sub ButtonArrangeHorizontal(control As IRibbonControl)

    Dim shp As Shape
    Dim leftWatermark, increment, leftValue, widthValue, rotationAdj As Single
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Select two or more shapes."
        Exit Sub
    End If

    leftWatermark = 12345678
    increment = InputBox("Enter horizontal gap in points." & vbNewLine & "(72 points = 1 inch; negative values apply)")
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.Rotation = 270 Or shp.Rotation = 90 Then
            ' Determine left value for rotated shapes
            If shp.Width <> shp.Height Then
                leftValue = shp.Left + (shp.Width - shp.Height) / 2
            Else
                leftValue = shp.Left
            End If
        Else
            leftValue = shp.Left
        End If
        
        If leftValue < leftWatermark Then
            leftWatermark = leftValue
        End If
    Next shp
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Enter a numeric gap."
            Exit Sub
        Else
            ' Determine effective width of shape for calculations
            If shp.Rotation = 270 Or shp.Rotation = 90 Then
                widthValue = shp.Height
                rotationAdj = (shp.Height - shp.Width) / 2
            Else
                widthValue = shp.Width
                rotationAdj = 0
            End If
            
            ' Arrange shapes
            shp.Left = leftWatermark + rotationAdj
            leftWatermark = leftWatermark + widthValue + increment
        End If
    Next shp
    
End Sub

Sub ButtonMeasureGap(control As IRibbonControl)

    Dim shp As Shape
    Dim shape1WasMeasured As Boolean
    Dim shape1T, shape1L, shape1H, shape1W As Single
    Dim shape2T, shape2L, shape2H, shape2W As Single
    Dim VGap, HGap As Single
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes."
        Exit Sub
    End If

    shape1WasMeasured = False
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Err <> 0 Then
            MsgBox "Select a shape."
        Else
            If shape1WasMeasured = False Then
                If shp.Rotation = 270 Or shp.Rotation = 90 Then
                    shape1T = shp.Top + (shp.Height - shp.Width) / 2
                    shape1L = shp.Left + (shp.Width - shp.Height) / 2
                    shape1H = shp.Width
                    shape1W = shp.Height
                Else
                    shape1T = shp.Top
                    shape1L = shp.Left
                    shape1H = shp.Height
                    shape1W = shp.Width
                End If
                shape1WasMeasured = True
            Else
                If shp.Rotation = 270 Or shp.Rotation = 90 Then
                    shape2T = shp.Top + (shp.Height - shp.Width) / 2
                    shape2L = shp.Left + (shp.Width - shp.Height) / 2
                    shape2H = shp.Width
                    shape2W = shp.Height
                Else
                    shape2T = shp.Top
                    shape2L = shp.Left
                    shape2H = shp.Height
                    shape2W = shp.Width
                End If
            End If
        End If
    Next shp
    
    If shape1T <= shape2T Then
        VGap = shape2T - shape1T - shape1H
        VGap = Round(VGap, 1)
    ElseIf shape2T < shape1T Then
        VGap = shape1T - shape2T - shape2H
        VGap = Round(VGap, 1)
    Else
        VGap = "--"
    End If
    
    If shape1L <= shape2L Then
        HGap = shape2L - shape1L - shape1W
        HGap = Round(HGap, 1)
    ElseIf shape2L <= shape1L Then
        HGap = shape1L - shape2L - shape2W
        HGap = Round(HGap, 1)
    Else
        HGap = "--"
    End If
    
    MsgBox "Vertical Gap = " & VGap & " Points" & vbNewLine & "Horizontal Gap = " & HGap & " Points"
    
End Sub