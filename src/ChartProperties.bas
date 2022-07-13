Attribute VB_Name = "ChartProperties"

Dim refPlotTop, refPlotLeft, refPlotHeight, refPlotWidth As Single

Sub ButtonChartCopy(control As IRibbonControl)
    
    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    
    testIfChart = shp.Chart.PlotArea.Top

    If Err <> 0 Then
        MsgBox "Select a chart."
    Else
        refPlotTop = shp.Chart.PlotArea.InsideTop
        refPlotLeft = shp.Chart.PlotArea.InsideLeft
        refPlotHeight = shp.Chart.PlotArea.InsideHeight
        refPlotWidth = shp.Chart.PlotArea.InsideWidth

        ' Paste copied properties onto original shape to remove clipping
        shp.Chart.PlotArea.InsideTop = refPlotTop
        shp.Chart.PlotArea.InsideLeft = refPlotLeft
        shp.Chart.PlotArea.InsideHeight = refPlotHeight
        shp.Chart.PlotArea.InsideWidth = refPlotWidth
    End If
    
End Sub

Sub ButtonChartPasteTop(control As IRibbonControl)
    
    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refPlotHeight = 0 And refPlotWidth = 0 Then
        MsgBox "Copy chart properties first."
        Exit Sub
    End If
    
    testIfChart = ActiveWindow.Selection.ShapeRange(1).Chart.PlotArea.Top

    If Err <> 0 Then
        MsgBox "Select a chart."
    Else
        For Each shp In ActiveWindow.Selection.ShapeRange
            shp.Chart.PlotArea.InsideTop = refPlotTop
        Next shp
    End If
End Sub

Sub ButtonChartPasteBottom(control As IRibbonControl)
    
    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refPlotHeight = 0 And refPlotWidth = 0 Then
        MsgBox "Copy chart properties first."
        Exit Sub
    End If
    
    testIfChart = ActiveWindow.Selection.ShapeRange(1).Chart.PlotArea.Top

    If Err <> 0 Then
        MsgBox "Select a chart."
    Else
        For Each shp In ActiveWindow.Selection.ShapeRange
            shp.Chart.PlotArea.InsideTop = refPlotTop + refPlotHeight - shp.Chart.PlotArea.InsideHeight
        Next shp
    End If
End Sub

Sub ButtonChartPasteLeft(control As IRibbonControl)
    
    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refPlotHeight = 0 And refPlotWidth = 0 Then
        MsgBox "Copy chart properties first."
        Exit Sub
    End If
    
    testIfChart = ActiveWindow.Selection.ShapeRange(1).Chart.PlotArea.Top

    If Err <> 0 Then
        MsgBox "Select a chart."
    Else
        For Each shp In ActiveWindow.Selection.ShapeRange
            shp.Chart.PlotArea.InsideLeft = refPlotLeft
        Next shp
    End If
End Sub

Sub ButtonChartPasteRight(control As IRibbonControl)
    
    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refPlotHeight = 0 And refPlotWidth = 0 Then
        MsgBox "Copy chart properties first."
        Exit Sub
    End If
    
    testIfChart = ActiveWindow.Selection.ShapeRange(1).Chart.PlotArea.Top

    If Err <> 0 Then
        MsgBox "Select a chart."
    Else
        For Each shp In ActiveWindow.Selection.ShapeRange
            shp.Chart.PlotArea.InsideLeft = refPlotLeft + refPlotWidth - shp.Chart.PlotArea.InsideWidth
        Next shp
    End If
End Sub

Sub ButtonChartPasteHeight(control As IRibbonControl)
    
    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refPlotHeight = 0 And refPlotWidth = 0 Then
        MsgBox "Copy chart properties first."
        Exit Sub
    End If
    
    testIfChart = ActiveWindow.Selection.ShapeRange(1).Chart.PlotArea.Top

    If Err <> 0 Then
        MsgBox "Select a chart."
    Else
        For Each shp In ActiveWindow.Selection.ShapeRange
            shp.Chart.PlotArea.InsideHeight = refPlotHeight
        Next shp
    End If
End Sub

Sub ButtonChartPasteWidth(control As IRibbonControl)
    
    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refPlotHeight = 0 And refPlotWidth = 0 Then
        MsgBox "Copy chart properties first."
        Exit Sub
    End If
    
    testIfChart = ActiveWindow.Selection.ShapeRange(1).Chart.PlotArea.Top

    If Err <> 0 Then
        MsgBox "Select a chart."
    Else
        For Each shp In ActiveWindow.Selection.ShapeRange
            shp.Chart.PlotArea.InsideWidth = refPlotWidth
        Next shp
    End If
End Sub

Sub ButtonChartPasteAll(control As IRibbonControl)
    
    Dim shp As Shape
    On Error Resume Next
    Err.Clear
    
    If refPlotHeight = 0 And refPlotWidth = 0 Then
        MsgBox "Copy chart properties first."
        Exit Sub
    End If
    
    testIfChart = ActiveWindow.Selection.ShapeRange(1).Chart.PlotArea.Top

    If Err <> 0 Then
        MsgBox "Select a chart."
    Else
        For Each shp In ActiveWindow.Selection.ShapeRange
            shp.Chart.PlotArea.InsideTop = refPlotTop
            shp.Chart.PlotArea.InsideLeft = refPlotLeft
            shp.Chart.PlotArea.InsideHeight = refPlotHeight
            shp.Chart.PlotArea.InsideWidth = refPlotWidth

            ' Repeat actions to overcome any clipping that may have happened in the previous step
            shp.Chart.PlotArea.InsideTop = refPlotTop
            shp.Chart.PlotArea.InsideLeft = refPlotLeft
            shp.Chart.PlotArea.InsideHeight = refPlotHeight
            shp.Chart.PlotArea.InsideWidth = refPlotWidth
        Next shp
    End If
End Sub
