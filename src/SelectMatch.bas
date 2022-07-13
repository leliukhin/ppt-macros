Attribute VB_Name = "SelectMatch"

Sub ButtonSelectTop(control As IRibbonControl)
    
    Dim referenceShape, shp As Shape
    Dim sld As Slide
    
    On Error Resume Next
    Err.Clear
    
    Set sld = Application.ActiveWindow.View.Slide
    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)

    If Err <> 0 Then
        MsgBox "Select a reference shape."
        Exit Sub
    End If
  
    For Each shp In sld.Shapes
        If shp.Top = referenceShape.Top Then
            shp.Select Replace:=False
        End If
    Next
    
End Sub

Sub ButtonSelectBottom(control As IRibbonControl)
    
    Dim referenceShape, shp As Shape
    Dim sld As Slide
    
    On Error Resume Next
    Err.Clear
    
    Set sld = Application.ActiveWindow.View.Slide
    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)

    If Err <> 0 Then
        MsgBox "Select a reference shape."
        Exit Sub
    End If

    For Each shp In sld.Shapes
        If shp.Top + shp.Height = referenceShape.Top + referenceShape.Height Then
            shp.Select Replace:=False
        End If
    Next
    
End Sub

Sub ButtonSelectLeft(control As IRibbonControl)
    
    Dim referenceShape, shp As Shape
    Dim sld As Slide
    
    On Error Resume Next
    Err.Clear
    
    Set sld = Application.ActiveWindow.View.Slide
    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)

    If Err <> 0 Then
        MsgBox "Select a reference shape."
        Exit Sub
    End If
  
    For Each shp In sld.Shapes
        If shp.Left = referenceShape.Left Then
            shp.Select Replace:=False
        End If
    Next
    
End Sub

Sub ButtonSelectRight(control As IRibbonControl)
    
    Dim referenceShape, shp As Shape
    Dim sld As Slide
    
    On Error Resume Next
    Err.Clear
    
    Set sld = Application.ActiveWindow.View.Slide
    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)

    If Err <> 0 Then
        MsgBox "Select a reference shape."
        Exit Sub
    End If
  
    For Each shp In sld.Shapes
        If shp.Left + shp.Width = referenceShape.Left + referenceShape.Width Then
            shp.Select Replace:=False
        End If
    Next
    
End Sub

Sub ButtonSelectHeight(control As IRibbonControl)
    
    Dim referenceShape, shp As Shape
    Dim sld As Slide
    
    On Error Resume Next
    Err.Clear
    
    Set sld = Application.ActiveWindow.View.Slide
    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)

    If Err <> 0 Then
        MsgBox "Select a reference shape."
        Exit Sub
    End If
  
    For Each shp In sld.Shapes
        If shp.Height = referenceShape.Height Then
            shp.Select Replace:=False
        End If
    Next
    
End Sub

Sub ButtonSelectWidth(control As IRibbonControl)
    
    Dim referenceShape, shp As Shape
    Dim sld As Slide
    
    On Error Resume Next
    Err.Clear
    
    Set sld = Application.ActiveWindow.View.Slide
    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)

    If Err <> 0 Then
        MsgBox "Select a reference shape."
        Exit Sub
    End If
    
    For Each shp In sld.Shapes
        If shp.Width = referenceShape.Width Then
            shp.Select Replace:=False
        End If
    Next
    
End Sub

Sub ButtonSelectSize(control As IRibbonControl)
    
    Dim referenceShape, shp As Shape
    Dim sld As Slide
    
    On Error Resume Next
    Err.Clear
    
    Set sld = Application.ActiveWindow.View.Slide
    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)

    If Err <> 0 Then
        MsgBox "Select a reference shape."
        Exit Sub
    End If
  
    For Each shp In sld.Shapes
        If shp.Height = referenceShape.Height And shp.Width = referenceShape.Width Then
            shp.Select Replace:=False
        End If
    Next
    
End Sub

Sub ButtonSelectLogoSize(control As IRibbonControl)
    
    Dim referenceShape, shp As Shape
    Dim sld As Slide
    
    On Error Resume Next
    Err.Clear
    
    Set sld = Application.ActiveWindow.View.Slide
    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)

    If Err <> 0 Then
        MsgBox "Select a reference shape."
        Exit Sub
    End If
    
    For Each shp In sld.Shapes
        If Round(shp.Height * shp.Width, 0) = Round(referenceShape.Height * referenceShape.Width, 0) Then
            shp.Select Replace:=False
        End If
    Next
    
End Sub

Sub ButtonSelectShapeType(control As IRibbonControl)
    
    Dim referenceShape, shp As Shape
    Dim sld As Slide
    
    On Error Resume Next
    Err.Clear
    
    Set sld = Application.ActiveWindow.View.Slide
    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)

    If Err <> 0 Then
        MsgBox "Select a reference shape."
        Exit Sub
    End If
    
    For Each shp In sld.Shapes
        If shp.AutoShapeType = referenceShape.AutoShapeType _
            And shp.HasTextFrame = referenceShape.HasTextFrame _
        Then
            shp.Select Replace:=False
        End If
    Next
    
End Sub

Sub ButtonSelectFill(control As IRibbonControl)
    
    Dim referenceShape, shp As Shape
    Dim sld As Slide
    
    On Error Resume Next
    Err.Clear
    
    Set sld = Application.ActiveWindow.View.Slide
    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)

    If Err <> 0 Then
        MsgBox "Select a reference shape."
        Exit Sub
    End If
  
    For Each shp In sld.Shapes
        If shp.Fill.Visible = msoTrue _
            And shp.Fill.ForeColor.RGB = referenceShape.Fill.ForeColor.RGB _
        Then
            shp.Select Replace:=False
        End If
    Next
    
End Sub

Sub ButtonSelectLine(control As IRibbonControl)
    
    Dim referenceShape, shp As Shape
    Dim sld As Slide
    
    On Error Resume Next
    Err.Clear
    
    Set sld = Application.ActiveWindow.View.Slide
    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)

    If Err <> 0 Then
        MsgBox "Select a reference shape."
        Exit Sub
    End If
  
    For Each shp In sld.Shapes
        If _
            shp.Line.Visible = msoTrue _
            And shp.Line.ForeColor.RGB = referenceShape.Line.ForeColor.RGB _
            And shp.Line.DashStyle = referenceShape.Line.DashStyle _
        Then
            shp.Select Replace:=False
        End If
    Next
    
End Sub
