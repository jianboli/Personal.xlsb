Attribute VB_Name = "QuickFormat"
Sub QuickFormat()
    Dim rng As Range
    Dim sh As Worksheet
    Set rng = Selection
    Set sh = Application.ActiveSheet
    
    Set rng = Intersect(rng, sh.UsedRange)
    If rng Is Nothing Then
        Set rng = sh.UsedRange
    ElseIf rng.Columns.Count = 0 Or rng.Rows.Count <= 1 Then
        Set rng = sh.UsedRange
    End If
    
    'auto width
    rng.Columns.EntireColumn.AutoFit
    rng.Font.Size = 10
    
    ' border
    With rng.borders(xlEdgeLeft)
        .LineStyle = xlContinuous
    End With
    With rng.borders(xlEdgeTop)
        .LineStyle = xlContinuous
    End With
    With rng.borders(xlEdgeBottom)
        .LineStyle = xlContinuous
    End With
    With rng.borders(xlEdgeRight)
        .LineStyle = xlContinuous
    End With
    With rng.borders(xlInsideVertical)
        .LineStyle = xlContinuous
    End With
    With rng.borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
    End With
    Selection.Font.Size = 10
    
    'format headers
    Dim header As Range
    Set header = rng.Rows(1)

    With header.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    header.Font.Bold = True
    With header.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    header.HorizontalAlignment = xlCenter

    rng.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    rng.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"

    ' check if it the column be percentage
    Dim i As Long, j As Long
    
    For j = 1 To rng.Columns.Count
        Dim possiblePercentage As Boolean
        possiblePercentage = True
        Dim possibleDate As Boolean
        possibleDate = True
        
        For i = 2 To rng.Rows.Count ' assume headers always
            If Not IsNumeric(rng.Cells(i, j).Value) Then
                possiblePercentage = False
                possibleDate = False
                GoTo EndOfI
            End If
        
            If Not IsEmpty(rng.Cells(i, j)) And Abs(rng.Cells(i, j).Value) > 10 Then
                possiblePercentage = False
            End If
            
            If Not IsEmpty(rng.Cells(i, j)) And (rng.Cells(i, j).Value < 29221 Or rng.Cells(i, j).Value > 54789) Then   '1980-1-1 to 2050-1-1
                possibleDate = False
            End If
        Next i
EndOfI:
    If (possiblePercentage) Then ' probably percentage value
        rng.Columns(j).NumberFormat = "0.00%"
    End If
    
    If (possibleDate) Then ' probably Date
        rng.Columns(j).NumberFormat = "mm/dd/yyyy"
    End If
    Next j
    
    'Replace null values with blank cell
    rng.Replace "[NULL]", ""
    rng.Replace "NULL", ""
    
    ' Application.Goto rng.Cells(1, 1), True
    ' With ActiveWindow
    '     If .FreezePanes Then .FreezePanes = False
    '     .SplitColumn = 0
    '     .SplitRow = 1
    '     .FreezePanes = True
    'End With
    
End Sub


