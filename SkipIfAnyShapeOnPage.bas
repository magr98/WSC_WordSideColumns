Attribute VB_Name = "SkipIfAnyShapeOnPage"
Sub AddSideColumnOnPagesWithoutColumns()
    Dim shp As Shape
    Dim newShp As Shape
    Dim totalPages As Long, currentPage As Long
    Dim foundShapeOnPage As Boolean
    
    Dim colWidth As Single, leftMargin As Single, topMargin As Single
    Dim bottomMargin As Single, pageHeight As Single
    
    colWidth = CentimetersToPoints(4)
    leftMargin = ActiveDocument.PageSetup.leftMargin
    topMargin = ActiveDocument.PageSetup.topMargin
    bottomMargin = ActiveDocument.PageSetup.bottomMargin
    pageHeight = ActiveDocument.PageSetup.pageHeight
    
    totalPages = ActiveDocument.ComputeStatistics(wdStatisticPages)
    
    For currentPage = 1 To totalPages
        Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=currentPage
        Selection.Collapse Direction:=wdCollapseStart
        
        foundShapeOnPage = False
        
        For Each shp In ActiveDocument.Shapes
            If shp.Anchor.Information(wdActiveEndPageNumber) = currentPage Then
                foundShapeOnPage = True
                Exit For
            End If
        Next shp
                
        If Not foundShapeOnPage Then
            Set newShp = ActiveDocument.Shapes.AddTextbox( _
                Orientation:=msoTextOrientationHorizontal, _
                Left:=leftMargin - colWidth - CentimetersToPoints(0.2), _
                Top:=topMargin, _
                Width:=colWidth, _
                Height:=pageHeight - topMargin - bottomMargin, _
                Anchor:=Selection.Range)
            
            With newShp
                .Line.Visible = msoFalse
                .Fill.Solid
                .Fill.BackColor.RGB = RGB(255, 255, 255)
                .Fill.Transparency = 1
                .TextFrame.TextRange.Text = ""
                .TextFrame.MarginLeft = 5
                .TextFrame.MarginRight = 5
                .TextFrame.MarginTop = 5
                .TextFrame.MarginBottom = 5
                .WrapFormat.Type = wdWrapBehind
                .RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
                .RelativeVerticalPosition = wdRelativeVerticalPositionPage
                .LockAnchor = True
                .LockAspectRatio = msoTrue
                .TextFrame.AutoSize = False
            End With
        End If
    Next currentPage
    
    MsgBox "Side columns added on pages without shapes anchored!"
End Sub


