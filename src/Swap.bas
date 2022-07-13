Attribute VB_Name = "Swap"

Sub ButtonSwapTL(control As IRibbonControl)

    Dim shape1, shape2 As Shape
    Dim shape1T, shape1L, shape1H, shape1W As Single
    Dim shape2T, shape2L, shape2H, shape2W As Single
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes to swap."
        Exit Sub
    End If

    Set shape1 = ActiveWindow.Selection.ShapeRange(1)
    Set shape2 = ActiveWindow.Selection.ShapeRange(2)

    'Copy starting properties of both shapes
    shape1T = shape1.Top
    shape1L = shape1.Left

    shape2T = shape2.Top
    shape2L = shape2.Left
      
    'Apply swap
    shape1.Top = shape2T
    shape1.Left = shape2L

    shape2.Top = shape1T
    shape2.Left = shape1L
    
End Sub

Sub ButtonSwapTR(control As IRibbonControl)

    Dim shape1, shape2 As Shape
    Dim shape1T, shape1L, shape1H, shape1W As Single
    Dim shape2T, shape2L, shape2H, shape2W As Single
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes to swap."
        Exit Sub
    End If

    Set shape1 = ActiveWindow.Selection.ShapeRange(1)
    Set shape2 = ActiveWindow.Selection.ShapeRange(2)

    'Copy starting properties of both shapes
    shape1T = shape1.Top
    shape1L = shape1.Left
    shape1W = shape1.Width

    shape2T = shape2.Top
    shape2L = shape2.Left
    shape2W = shape2.Width
    
    'Apply swap
    shape1.Top = shape2T
    shape1.Left = shape2L + shape2W - shape1W
    
    shape2.Top = shape1T
    shape2.Left = shape1L + shape1W - shape2W

End Sub

Sub ButtonSwapBL(control As IRibbonControl)

    Dim shape1, shape2 As Shape
    Dim shape1T, shape1L, shape1H, shape1W As Single
    Dim shape2T, shape2L, shape2H, shape2W As Single
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes to swap."
        Exit Sub
    End If

    Set shape1 = ActiveWindow.Selection.ShapeRange(1)
    Set shape2 = ActiveWindow.Selection.ShapeRange(2)

    'Copy starting properties of both shapes
    shape1T = shape1.Top
    shape1L = shape1.Left
    shape1H = shape1.Height

    shape2T = shape2.Top
    shape2L = shape2.Left
    shape2H = shape2.Height
    
    'Apply swap
    shape1.Top = shape2T + shape2H - shape1H
    shape1.Left = shape2L

    shape2.Top = shape1T + shape1H - shape2H
    shape2.Left = shape1L
    
End Sub

Sub ButtonSwapBR(control As IRibbonControl)

    Dim shape1, shape2 As Shape
    Dim shape1T, shape1L, shape1H, shape1W As Single
    Dim shape2T, shape2L, shape2H, shape2W As Single
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes to swap."
        Exit Sub
    End If

    Set shape1 = ActiveWindow.Selection.ShapeRange(1)
    Set shape2 = ActiveWindow.Selection.ShapeRange(2)

    'Copy starting properties of both shapes
    shape1T = shape1.Top
    shape1L = shape1.Left
    shape1H = shape1.Height
    shape1W = shape1.Width

    shape2T = shape2.Top
    shape2L = shape2.Left
    shape2H = shape2.Height
    shape2W = shape2.Width
    
    'Apply swap
    shape1.Top = shape2T + shape2H - shape1H
    shape1.Left = shape2L + shape2W - shape1W
    
    shape2.Top = shape1T + shape1H - shape2H
    shape2.Left = shape1L + shape1W - shape2W
    
End Sub

Sub ButtonSwapCM(control As IRibbonControl)

    Dim shape1, shape2 As Shape
    Dim shape1T, shape1L, shape1H, shape1W As Single
    Dim shape2T, shape2L, shape2H, shape2W As Single
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes to swap."
        Exit Sub
    End If

    Set shape1 = ActiveWindow.Selection.ShapeRange(1)
    Set shape2 = ActiveWindow.Selection.ShapeRange(2)

    'Copy starting properties of both shapes
    shape1T = shape1.Top
    shape1L = shape1.Left
    shape1H = shape1.Height
    shape1W = shape1.Width

    shape2T = shape2.Top
    shape2L = shape2.Left
    shape2H = shape2.Height
    shape2W = shape2.Width
    
    'Apply swap
    shape1.Top = shape2T + shape2H / 2 - shape1H / 2
    shape1.Left = shape2L + shape2W / 2 - shape1W / 2

    shape2.Top = shape1T + shape1H / 2 - shape2H / 2
    shape2.Left = shape1L + shape1W / 2 - shape2W / 2         
    
End Sub

Sub ButtonSwapCC(control As IRibbonControl)

    Dim shape1, shape2 As Shape
    Dim shape1T, shape1L, shape1H, shape1W As Single
    Dim shape2T, shape2L, shape2H, shape2W As Single
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes to swap."
        Exit Sub
    End If

    Set shape1 = ActiveWindow.Selection.ShapeRange(1)
    Set shape2 = ActiveWindow.Selection.ShapeRange(2)

    'Copy starting properties of both shapes
    shape1L = shape1.Left
    shape1W = shape1.Width

    shape2L = shape2.Left
    shape2W = shape2.Width
    
    'Apply swap
    shape1.Left = shape2L + shape2W / 2 - shape1W / 2

    shape2.Left = shape1L + shape1W / 2 - shape2W / 2
    
End Sub

Sub ButtonSwapMM(control As IRibbonControl)

    Dim shape1, shape2 As Shape
    Dim shape1T, shape1L, shape1H, shape1W As Single
    Dim shape2T, shape2L, shape2H, shape2W As Single
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes to swap."
        Exit Sub
    End If

    Set shape1 = ActiveWindow.Selection.ShapeRange(1)
    Set shape2 = ActiveWindow.Selection.ShapeRange(2)

    'Copy starting properties of both shapes
    shape1T = shape1.Top
    shape1H = shape1.Height

    shape2T = shape2.Top
    shape2H = shape2.Height
    
    'Apply swap
    shape1.Top = shape2T + shape2H / 2 - shape1H / 2

    shape2.Top = shape1T + shape1H / 2 - shape2H / 2
    
End Sub

Sub ButtonSwapCT(control As IRibbonControl)

    Dim shape1, shape2 As Shape
    Dim shape1T, shape1L, shape1H, shape1W As Single
    Dim shape2T, shape2L, shape2H, shape2W As Single
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes to swap."
        Exit Sub
    End If

    Set shape1 = ActiveWindow.Selection.ShapeRange(1)
    Set shape2 = ActiveWindow.Selection.ShapeRange(2)

    'Copy starting properties of both shapes
    shape1T = shape1.Top
    shape1L = shape1.Left
    shape1W = shape1.Width

    shape2T = shape2.Top
    shape2L = shape2.Left
    shape2W = shape2.Width
    
    'Apply swap
    shape1.Top = shape2T
    shape1.Left = shape2L + shape2W / 2 - shape1W / 2
    
    shape2.Top = shape1T
    shape2.Left = shape1L + shape1W / 2 - shape2W / 2
    
End Sub

Sub ButtonSwapCB(control As IRibbonControl)

    Dim shape1, shape2 As Shape
    Dim shape1T, shape1L, shape1H, shape1W As Single
    Dim shape2T, shape2L, shape2H, shape2W As Single
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes to swap."
        Exit Sub
    End If

    Set shape1 = ActiveWindow.Selection.ShapeRange(1)
    Set shape2 = ActiveWindow.Selection.ShapeRange(2)

    'Copy starting properties of both shapes
    shape1T = shape1.Top
    shape1L = shape1.Left
    shape1H = shape1.Height
    shape1W = shape1.Width

    shape2T = shape2.Top
    shape2L = shape2.Left
    shape2H = shape2.Height
    shape2W = shape2.Width
    
    'Apply swap
    shape1.Top = shape2T + shape2H - shape1H
    shape1.Left = shape2L + shape2W / 2 - shape1W / 2

    shape2.Top = shape1T + shape1H - shape2H
    shape2.Left = shape1L + shape1W / 2 - shape2W / 2
    
End Sub

Sub ButtonSwapML(control As IRibbonControl)

    Dim shape1, shape2 As Shape
    Dim shape1T, shape1L, shape1H, shape1W As Single
    Dim shape2T, shape2L, shape2H, shape2W As Single
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes to swap."
        Exit Sub
    End If

    Set shape1 = ActiveWindow.Selection.ShapeRange(1)
    Set shape2 = ActiveWindow.Selection.ShapeRange(2)

    'Copy starting properties of both shapes
    shape1T = shape1.Top
    shape1L = shape1.Left
    shape1H = shape1.Height

    shape2T = shape2.Top
    shape2L = shape2.Left
    shape2H = shape2.Height
    
    'Apply swap
    shape1.Top = shape2T + shape2H / 2 - shape1H / 2
    shape1.Left = shape2L

    shape2.Top = shape1T + shape1H / 2 - shape2H / 2
    shape2.Left = shape1L
    
End Sub

Sub ButtonSwapMR(control As IRibbonControl)

    Dim shape1, shape2 As Shape
    Dim shape1T, shape1L, shape1H, shape1W As Single
    Dim shape2T, shape2L, shape2H, shape2W As Single
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes to swap."
        Exit Sub
    End If

    Set shape1 = ActiveWindow.Selection.ShapeRange(1)
    Set shape2 = ActiveWindow.Selection.ShapeRange(2)

    'Copy starting properties of both shapes
    shape1T = shape1.Top
    shape1L = shape1.Left
    shape1H = shape1.Height
    shape1W = shape1.Width

    shape2T = shape2.Top
    shape2L = shape2.Left
    shape2H = shape2.Height
    shape2W = shape2.Width
    
    'Apply swap
    shape1.Top = shape2T + shape2H / 2 - shape1H / 2
    shape1.Left = shape2L + shape2W - shape1W
    
    shape2.Top = shape1T + shape1H / 2 - shape2H / 2
    shape2.Left = shape1L + shape1W - shape2W
    
End Sub

Sub ButtonSwapOH(control As IRibbonControl)

    Dim shape1, shape2 As Shape
    Dim shape1T, shape1L, shape1H, shape1W As Single
    Dim shape2T, shape2L, shape2H, shape2W As Single
    Dim shape1OnLeft, shape2OnLeft As Integer
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes to swap."
        Exit Sub
    End If

    Set shape1 = ActiveWindow.Selection.ShapeRange(1)
    Set shape2 = ActiveWindow.Selection.ShapeRange(2)

    'Copy starting properties of both shapes
    shape1L = shape1.Left
    shape1W = shape1.Width

    shape2L = shape2.Left
    shape2W = shape2.Width
    
    If shape1L > shape2L Then
        shape1OnLeft = 0
        shape2OnLeft = 1
    Else
        shape1OnLeft = 1
        shape2OnLeft = 0
    End If
    
    'Apply swap
    shape1.Left = shape2L + shape1OnLeft * shape2W - shape1OnLeft * shape1W

    shape2.Left = shape1L + shape2OnLeft * shape1W - shape2OnLeft * shape2W
    
End Sub

Sub ButtonSwapOV(control As IRibbonControl)

    Dim shape1, shape2 As Shape
    Dim shape1T, shape1L, shape1H, shape1W As Single
    Dim shape2T, shape2L, shape2H, shape2W As Single
    Dim shape1OnTop, shape2OnTop As Integer
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes to swap."
        Exit Sub
    End If

    Set shape1 = ActiveWindow.Selection.ShapeRange(1)
    Set shape2 = ActiveWindow.Selection.ShapeRange(2)

    'Copy starting properties of both shapes
    shape1T = shape1.Top
    shape1H = shape1.Height

    shape2T = shape2.Top
    shape2H = shape2.Height
    
    If shape1T > shape2T Then
        shape1OnTop = 0
        shape2OnTop = 1
    Else
        shape1OnTop = 1
        shape2OnTop = 0
    End If
    
    'Apply swap
    shape1.Top = shape2T + shape1OnTop * shape2H - shape1OnTop * shape1H

    shape2.Top = shape1T + shape2OnTop * shape1H - shape2OnTop * shape2H
    
End Sub
