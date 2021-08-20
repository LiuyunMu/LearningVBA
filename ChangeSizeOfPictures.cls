Sub 图片排版()
'
' 图片排版 宏
'
'
Dim i As Integer

    For i = 1 To ActiveDocument.InlineShapes.Count
    
    With ActiveDocument.InlineShapes(i)
    
    .Height = 238 * 0.7
    
    .Width = 315 * 0.8
    
    .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
    
    .Borders(wdBorderTop).LineWidth = wdLineWidth050pt
    
    .Borders(wdBorderTop).Color = wdColorAutomatic
    
    End With
    
    Next i
End Sub
