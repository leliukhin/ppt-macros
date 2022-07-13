VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufAddMatrix 
   Caption         =   "Add Grid"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3615
   OleObjectBlob   =   "ufAddMatrix.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufAddMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ufAddMatrix_Initialize()
    Dim ufMatrixDefaultColumns As Integer
    Dim ufMatrixDefaultRows As Integer
    Dim ufMatrixDefaultSpacing As Integer
    
    ufMatrixDefaultColumns = 3
    ufMatrixDefaultRows = 3
    ufMatrixDefaultSpacing = 5
    
    ufMatrixCBColumnHeaders.Value = False
    ufMatrixCBRowHeaders.Value = False
    ufMatrixSpinColumns.Value = ufMatrixDefaultColumns
    ufMatrixSpinRows.Value = ufMatrixDefaultRows
    ufMatrixSpinSpacing.Value = ufMatrixDefaultSpacing
    ufMatrixTextColumns.Text = ufMatrixDefaultColumns
    ufMatrixTextRows.Text = ufMatrixDefaultRows
    ufMatrixTextSpacing.Text = ufMatrixDefaultSpacing
End Sub

Private Sub ufMatrixCancelButton_Click()
    Unload Me
End Sub

Private Sub ufMatrixSpinColumns_Change()
    ufMatrixTextColumns.Text = ufMatrixSpinColumns.Value
End Sub

Private Sub ufMatrixTextColumns_Change()
    
    On Error Resume Next
    newColVal = Val(ufMatrixTextColumns.Text)
    
    If CInt(newColVal) = newColVal Then
        If newColVal >= ufMatrixSpinColumns.Min And newColVal <= ufMatrixSpinColumns.Max Then
            ufMatrixSpinColumns.Value = newColVal
        Else
            MsgBox "Please enter a value between " & ufMatrixSpinColumns.Min & " and " & ufMatrixSpinColumns.Max
            ufMatrixSpinColumns.Value = ufMatrixDefaultColumns
            ufMatrixTextColumns.Text = ufMatrixSpinColumns.Value
        End If
    Else
        MsgBox "Please enter a value between " & ufMatrixSpinColumns.Min & " and " & ufMatrixSpinColumns.Max
        ufMatrixSpinColumns.Value = ufMatrixDefaultColumns
        ufMatrixTextColumns.Text = ufMatrixSpinColumns.Value
    End If
End Sub

Private Sub ufMatrixSpinRows_Change()
    ufMatrixTextRows.Text = ufMatrixSpinRows.Value
End Sub

Private Sub ufMatrixTextRows_Change()
    
    On Error Resume Next
    newRowVal = Val(ufMatrixTextRows.Text)
    
    If CInt(newRowVal) = newRowVal Then
        If newRowVal >= ufMatrixSpinRows.Min And newRowVal <= ufMatrixSpinRows.Max Then
            ufMatrixSpinRows.Value = newRowVal
        Else
            MsgBox "Please enter a value between " & ufMatrixSpinRows.Min & " and " & ufMatrixSpinRows.Max
            ufMatrixSpinRows.Value = ufMatrixDefaultRows
            ufMatrixTextRows.Text = ufMatrixSpinRows.Value
        End If
    Else
        MsgBox "Please enter a value between " & ufMatrixSpinRows.Min & " and " & ufMatrixSpinRows.Max
        ufMatrixSpinRows.Value = ufMatrixDefaultRows
        ufMatrixTextRows.Text = ufMatrixSpinRows.Value
    End If
End Sub

Private Sub ufMatrixSpinSpacing_Change()
    ufMatrixTextSpacing.Text = ufMatrixSpinSpacing.Value
End Sub

Private Sub ufMatrixTextSpacing_Change()
    
    On Error Resume Next
    newSpacingVal = Val(ufMatrixTextSpacing.Text)
    
    If CInt(newSpacingVal) = newSpacingVal Then
        If newSpacingVal >= ufMatrixSpinSpacing.Min And newSpacingVal <= ufMatrixSpinSpacing.Max Then
            ufMatrixSpinSpacing.Value = newSpacingVal
        Else
            MsgBox "Please enter a value between " & ufMatrixSpinSpacing.Min & " and " & ufMatrixSpinSpacing.Max
            ufMatrixTextSpacing.Text = ufMatrixDefaultSpacing
            ufMatrixTextSpacing.Text = ufMatrixSpinSpacing.Value
        End If
    Else
        MsgBox "Please enter a value between " & ufMatrixSpinSpacing.Min & " and " & ufMatrixSpinSpacing.Max
        ufMatrixTextSpacing.Text = ufMatrixDefaultSpacing
        ufMatrixTextSpacing.Text = ufMatrixSpinSpacing.Value
    End If
End Sub

Private Sub ufMatrixOKButton_Click()
    
    ' Pick up outer dimensions of the grid
    
    Dim oshp As Shape
    On Error Resume Next
    Err.Clear
    Set oshp = ActiveWindow.Selection.ShapeRange(1)
    
    If Err <> 0 Then
        MsgBox "Select a container shape."
        Exit Sub
    Else
        gridL = oshp.Left
        gridT = oshp.Top
        gridW = oshp.Width
        gridH = oshp.Height
    End If
    
    ' Determine which headers are used
    
    Dim CHeaders As Integer
    Dim RHeaders As Integer
    
    If ufMatrixCBColumnHeaders.Value = True Then
        CHeaders = 1
    Else
        CHeaders = 0
    End If
    
    If ufMatrixCBRowHeaders.Value = True Then
        RHeaders = 1
    Else
        RHeaders = 0
    End If
    
    ' Establish basic grid dimensions
    
    Dim NCols As Integer
    Dim NRows As Integer
    Dim GGap As Integer
    
    NCols = ufMatrixSpinColumns.Value
    NRows = ufMatrixSpinRows.Value
    GGap = ufMatrixSpinSpacing.Value
    
    ' Establish header proportions
    
    Dim ColHeader As Single
    Dim RowHeader As Single
    
    If CHeaders = 1 Then
        ColHeader = 0.5 ' Each column header will be 0.5 the height of a cell
    Else
        ColHeader = 0
    End If
    
    If RHeaders = 1 Then
        RowHeader = 1
    Else
        RowHeader = 0
    End If
    
    ' Establish graphic properties of grid
    
    RowHeaderColor = ActivePresentation.Designs(1).SlideMaster.Theme.ThemeColorScheme.Colors(msoThemeAccent1).RGB
    ColHeaderColor = ActivePresentation.Designs(1).SlideMaster.Theme.ThemeColorScheme.Colors(msoThemeAccent2).RGB
    InnerCellColor = RGB(173, 181, 189)
    LineWeight = 0.5
    
    ' Calculate the height and width of each cell
    
    CellHeight = (gridH - NRows * GGap - (CHeaders - 1) * GGap) / (NRows + ColHeader)
    CellWidth = (gridW - NCols * GGap - (RHeaders - 1) * GGap) / (NCols + RowHeader)
    
    ' Draw column headers
    
    Dim GSlide As Slide
    Set GSlide = Application.ActiveWindow.View.Slide
    
    If CHeaders = 1 Then
        For i = 1 To NCols
            Dim ColHeaderShp As Shape
            
            Set ColHeaderShp = GSlide.Shapes.AddShape(msoShapeRectangle, _
                Left:=gridL + CellWidth * RowHeader + CellWidth * (i - 1) + GGap * (i - (1 - RHeaders)), _
                Top:=gridT, _
                Width:=CellWidth, _
                Height:=CellHeight * ColHeader)
            
            ColHeaderShp.Fill.ForeColor.RGB = ColHeaderColor
            ColHeaderShp.Line.ForeColor.RGB = ColHeaderColor
            ColHeaderShp.Line.Weight = LineWeight
            ColHeaderShp.TextFrame.TextRange.Font.Bold = msoTrue
        Next i
    End If
    
    ' Draw row headers
    
    If RHeaders = 1 Then
        For i = 1 To NRows
            Dim RowHeaderShp As Shape
            
            Set RowHeaderShp = GSlide.Shapes.AddShape(msoShapeRectangle, _
                Left:=gridL, _
                Top:=gridT + CellHeight * ColHeader + CellHeight * (i - 1) + GGap * (i - (1 - CHeaders)), _
                Width:=CellWidth * RowHeader, _
                Height:=CellHeight)
            
            RowHeaderShp.Fill.ForeColor.RGB = RowHeaderColor
            RowHeaderShp.Line.ForeColor.RGB = RowHeaderColor
            RowHeaderShp.Line.Weight = LineWeight
            RowHeaderShp.TextFrame.TextRange.Font.Bold = msoTrue
        Next i
    End If
    
    ' Draw inner cells
    
    For i = 1 To NRows
        For j = 1 To NCols
            Dim CellShp As Shape
            
            Set CellShp = GSlide.Shapes.AddShape(msoShapeRectangle, _
                Left:=gridL + CellWidth * RowHeader + CellWidth * (j - 1) + GGap * (j - (1 - RHeaders)), _
                Top:=gridT + CellHeight * ColHeader + CellHeight * (i - 1) + GGap * (i - (1 - CHeaders)), _
                Width:=CellWidth, _
                Height:=CellHeight)
            
            CellShp.Fill.Visible = msoFalse
            CellShp.Line.ForeColor.RGB = InnerCellColor
            CellShp.Line.Weight = LineWeight
            
            CellShp.TextFrame.marginTop = 6
            CellShp.TextFrame.marginBottom = 6
            CellShp.TextFrame.marginLeft = 6
            CellShp.TextFrame.marginRight = 6
            CellShp.TextFrame.TextRange.Font.Color = ActivePresentation.Designs(1).SlideMaster.Theme.ThemeColorScheme.Colors(msoThemeDark1).RGB
            CellShp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
        Next j
    Next i
    
    oshp.Delete
    
    Unload Me
End Sub
