Sub toggleHighlighting()
'
' toggleHighlighting Macro
'
'
            
    If Selection.Range.HighlightColorIndex = False Then
        Options.DefaultHighlightColorIndex = wdYellow
        Selection.Range.HighlightColorIndex = wdYellow
    Else
        Options.DefaultHighlightColorIndex = wdYellow
        Selection.Range.HighlightColorIndex = wdNoHighlight
    End If

End Sub
