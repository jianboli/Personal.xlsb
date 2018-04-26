VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsertPerformanceTable 
   Caption         =   "UserForm1"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5655
   OleObjectBlob   =   "InsertPerformanceTable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InsertPerformanceTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    InsertTable
End Sub

Sub InsertTable()

    Selection.NumberFormat = "0"
    ActiveCell.FormulaR1C1 = "1"
    Range("V4").Select
    Selection.AutoFill Destination:=Range("V4:V15"), Type:=xlFillSeries
    Range("V4:V15").Select
    Range("X3").Select
    ActiveCell.FormulaR1C1 = "2011"
    Range("X3").Select
    Selection.AutoFill Destination:=Range("W3:X3"), Type:=xlFillSeries
    Range("W3:X3").Select
    Range("X3").Select
    Selection.AutoFill Destination:=Range("X3:AE3"), Type:=xlFillSeries
    Range("X3:AE3").Select
    Range("W4").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMPRODUCT((MONTH(R3C1:R141C1)=RC22)*(YEAR(R3C1:R141C1)=R3C)*(R3C3:R141C3))"
    Range("W4").Select
    Selection.AutoFill Destination:=Range("W4:W15"), Type:=xlFillDefault
    Range("W4:W15").Select
    Selection.AutoFill Destination:=Range("W4:AE15"), Type:=xlFillDefault
    Range("W4:AE15").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    Range("V16").Select
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveCell.FormulaR1C1 = "Total"
    Range("V17").Select
    ActiveCell.FormulaR1C1 = "Cumulative"
    Range("W16").Select
    Selection.FormulaArray = "=PRODUCT(1+R[-12]C:R[-1]C)"
    Selection.FormulaArray = "=PRODUCT(1+R[-12]C:R[-1]C)-1"
    Selection.AutoFill Destination:=Range("W16:AE16"), Type:=xlFillDefault
    Range("W16:AE16").Select
    Range("W17").Select
    ActiveCell.FormulaR1C1 = "1+"
    Range("W17").Select
    ActiveCell.FormulaR1C1 = "=(R[-1]C[-1]+1)*(1+R[-1]C)-1"
    Range("W17").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C"
    Range("X17").Select
    Selection.Style = "Percent"
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.NumberFormat = "0.00%"
    ActiveCell.FormulaR1C1 = "=(1+RC[-1])*(1+R[-1]C)-1"
    Range("X17").Select
    Selection.AutoFill Destination:=Range("X17:AE17"), Type:=xlFillDefault
    Range("X17:AE17").Select
    Range("W16:AE17").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    Columns("V:V").EntireColumn.AutoFit
    Range("V3:V15").Select
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
    End With
    Range("W3:AE3").Select
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
    End With
    Range("V2:AE2").Select
    Selection.borders(xlDiagonalDown).LineStyle = xlNone
    Selection.borders(xlDiagonalUp).LineStyle = xlNone
    Selection.borders(xlEdgeLeft).LineStyle = xlNone
    Selection.borders(xlEdgeTop).LineStyle = xlNone
    With Selection.borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.borders(xlEdgeRight).LineStyle = xlNone
    Selection.borders(xlInsideVertical).LineStyle = xlNone
    Selection.borders(xlInsideHorizontal).LineStyle = xlNone
    Range("V15:AE15").Select
    Selection.borders(xlDiagonalDown).LineStyle = xlNone
    Selection.borders(xlDiagonalUp).LineStyle = xlNone
    Selection.borders(xlEdgeLeft).LineStyle = xlNone
    Selection.borders(xlEdgeTop).LineStyle = xlNone
    With Selection.borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.borders(xlEdgeRight).LineStyle = xlNone
    Selection.borders(xlInsideVertical).LineStyle = xlNone
    Selection.borders(xlInsideHorizontal).LineStyle = xlNone
    Range("V17:AE17").Select
    Selection.borders(xlDiagonalDown).LineStyle = xlNone
    Selection.borders(xlDiagonalUp).LineStyle = xlNone
    Selection.borders(xlEdgeLeft).LineStyle = xlNone
    Selection.borders(xlEdgeTop).LineStyle = xlNone
    With Selection.borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.borders(xlEdgeRight).LineStyle = xlNone
    Selection.borders(xlInsideVertical).LineStyle = xlNone
    Selection.borders(xlInsideHorizontal).LineStyle = xlNone
    Range("V16:V17").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
    End With
    Range("W16:AE17").Select
    Selection.Font.Bold = True
    Range("W4:AE17").Select
    Selection.Font.Size = 10
    Range("V3:AE17").Select
    With Selection.Font
        .Name = "Arial"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Font.Size = 10
    Selection.Font.Size = 9
    Range("W4:AE15").Select
    Selection.NumberFormat = "0.00%;0.00%;"
    Selection.NumberFormat = "0.00%;-0.00%;"
    Range("V16:V17").Select
    Selection.Font.Bold = True
    Range("V4:V17").Select
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.349986266670736
    End With
    Range("W3:AE3").Select
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.349986266670736
    End With
    Range("V4:V15").Select
    Selection.Font.Bold = True
    Range("W3:AE3").Select
    Selection.Font.Bold = True
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
    End With
    Range("V4:V15").Select
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
    End With
    Range("V22").Select

End Sub
Private Sub RefEditDataRange_Change()
   Dim rng As Range
   If Not IsEmpty(RefEditDataRange.Value) Then
        
   End
End Sub

Private Sub UserForm_Initialize()
    Dim rng As Range
    If TypeName(Application.Selection) = "Range" Then
        Set rng = Application.Selection
        If rng.Count > 1 Then
            RefEditDataRange.Value = rng.Address
        End If
    End If
End Sub


