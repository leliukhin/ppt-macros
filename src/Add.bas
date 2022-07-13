Attribute VB_Name = "Add"

Sub ButtonDrawConnector(control As IRibbonControl)

    Dim sld As Slide
    Dim shp1, shp2, conn As Shape
    
    On Error Resume Next
    Err.Clear
    
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Select two shapes to connect."
        Exit Sub
    End If
    
    Set sld = Application.ActiveWindow.View.Slide
    
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    Set shp2 = ActiveWindow.Selection.ShapeRange(2)
    
    If shp1.Type = msoLine Or shp2.Type = msoLine Then
        MsgBox "Cannot draw connectors to a line."
        Exit Sub
    End If
    
    ' Straight connector if shapes are centered vertically or horizontally
    If _
        shp1.Top + shp1.Height / 2 = shp2.Top + shp2.Height / 2 _
        Or shp1.Left + shp1.Width / 2 = shp2.Left + shp2.Width / 2 _
    Then
        Set conn = sld.Shapes.AddConnector(msoConnectorStraight, 1, 1, 1, 1)
    Else
        Set conn = sld.Shapes.AddConnector(msoConnectorElbow, 1, 1, 1, 1)
    End If
    
    conn.ConnectorFormat.BeginConnect shp1, 1
    conn.ConnectorFormat.EndConnect shp2, 1

    conn.RerouteConnections
    
End Sub

Sub ButtonAddNTD(control As IRibbonControl)

    Dim sld As Slide
    Dim shp As Shape
    Dim ntdWidth As Integer
    Dim shpLeft As Single
    On Error Resume Next
    Err.Clear

    ntdWidth = 150
    shpLeft = ActivePresentation.PageSetup.SlideWidth - ntdWidth
    
    Set sld = Application.ActiveWindow.View.Slide
    Set shp = sld.Shapes.AddShape(Type:=msoShapeRectangle, _
        Left:=shpLeft, Top:=0, Width:=ntdWidth, Height:=70)
    shp.Fill.ForeColor.RGB = vbYellow
    shp.Line.Visible = False
    shp.TextFrame.marginTop = 6
    shp.TextFrame.marginBottom = 6
    shp.TextFrame.marginLeft = 6
    shp.TextFrame.marginRight = 6
    shp.TextFrame.TextRange.Font.Size = 11
    shp.TextFrame.TextRange.Font.Bold = True
    shp.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    shp.Select
    
End Sub

Sub ButtonUFAddMatrix()
    
    Dim referenceShape As Shape

    On Error Resume Next
    Err.Clear

    Set referenceShape = ActiveWindow.Selection.ShapeRange(1)
    
    If Err <> 0 Then
        MsgBox "Select a reference shape."
        Exit Sub
    End If
    
    ufAddMatrix.Show

End Sub