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
Option Explicit
Const ExpandableMaxRowNum As Integer = 10000

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    If RefEditDates.Value = "" Or RefEditMonthlyPnl.Value = "" Or RefEditInsertAt.Value = "" Then
        MsgBox "Please select all the ranges", vbCritical
        Exit Sub
    End If
    InsertTable
End Sub

Private Sub RefEditDataRange_Change()
   Dim rng As Range
   On Error GoTo ErrHandler
   If Not IsEmpty(RefEditDataRange.Value) Then
        Set rng = Range(RefEditDataRange.Value)
        
        If Not Application.WorksheetFunction.IsNumber(rng.Cells(1, 2)) Then
            Set rng = rng.Offset(1, 0).Resize(rng.Rows.Count - 1)
        End If
        
        If rng.Columns.Count = 2 Then
            RefEditDates.Value = rng.Columns(1).Address
            RefEditMonthlyPnl.Value = rng.Columns(2).Address
        End If
   End If
ErrHandler:
End Sub



Private Sub UserForm_Initialize()
    Dim rng As Range
    Dim sh As Worksheet
    
    Set sh = Application.ActiveSheet
    If TypeName(Application.Selection) = "Range" Then
        Set rng = Intersect(Application.Selection, sh.UsedRange)

        If rng.Count > 1 Then
            RefEditDataRange.Value = rng.Address
        End If
    End If
End Sub


Sub InsertTable()
    ' get the information
    Dim dateRng As Range
    Dim pnlRng As Range
    Dim inputRng As Range
    
    Set dateRng = Range(RefEditDates.Value)
    Set pnlRng = Range(RefEditMonthlyPnl.Value)
    ' Preprocess the dates
    Dim startDate As Date
    Dim endDate As Date
    
    On Error GoTo DateError
    startDate = Application.WorksheetFunction.Min(dateRng)
    endDate = Application.WorksheetFunction.Max(dateRng)
    On Error GoTo 0
    
    Set inputRng = Range(RefEditInsertAt).Cells(1, 1)
    ' Months
    Dim i As Integer
    For i = 1 To 12
        inputRng.Offset(i, 0) = i
    Next
    Dim totalYear As Integer
    totalYear = Year(endDate) - Year(startDate) + 1
    ' Years
    Dim j As Integer
    For j = 0 To totalYear - 1
        inputRng.Offset(0, j + 1) = j + Year(startDate)
    Next

    ' P&L Formulae
    Dim formulaStr As String
    Dim startRow As Integer, endRow As Integer, dateCol As Integer, pnlCol As Integer, monthCol As Integer, yearRow As Integer
    startRow = dateRng.Cells(1, 1).Row
    endRow = dateRng.Cells(dateRng.Rows.Count, 1).Row
    If CheckBoxExpandable.Value And endRow < ExpandableMaxRowNum Then
        endRow = ExpandableMaxRowNum
    End If
    pnlCol = pnlRng.Column
    dateCol = dateRng.Cells(1, 1).Column
    monthCol = inputRng.Column
    yearRow = inputRng.Row
    
    formulaStr = "=SUMPRODUCT((MONTH(R" & startRow & "C" & dateCol & ":R" & endRow & "C" & dateCol & ")=RC" & monthCol & ")" & _
                             "*(YEAR(R" & startRow & "C" & dateCol & ":R" & endRow & "C" & dateCol & ")=R" & yearRow & "C)" & _
                                  "*(R" & startRow & "C" & pnlCol & ":R" & endRow & "C" & pnlCol & "))"
    
    Dim contentRng As Range
    Set contentRng = Range(inputRng.Offset(1, 1), inputRng.Offset(13, totalYear))
    contentRng.FormulaR1C1 = formulaStr
    
    ' YTD Summary
    inputRng.Offset(13, 0).FormulaR1C1 = "Total"
 
    Set contentRng = Range(inputRng.Offset(13, 1), inputRng.Offset(13, totalYear))
    contentRng.FormulaArray = "=PRODUCT(1+R[-12]C:R[-1]C)-1"
    
    ' Cumulative return
    inputRng.Offset(14, 0).FormulaR1C1 = "Cumulative"
    inputRng.Offset(14, 1).FormulaR1C1 = "=R[-1]C"
    Set contentRng = Range(inputRng.Offset(14, 2), inputRng.Offset(14, totalYear))
    contentRng.FormulaR1C1 = "=(1+RC[-1])*(1+R[-1]C)-1"
    
    Range(inputRng.Offset(1, 1), inputRng.Offset(14, totalYear)).NumberFormat = "0.00%"
   
    ' formatting the headers
    With Range(inputRng, inputRng.Offset(13, 0)).Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
    End With
    
    With Range(inputRng, inputRng.Offset(0, totalYear)).Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
    End With
    
    ' format the border
    With Range(inputRng, inputRng.Offset(0, totalYear)).borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Range(inputRng.Offset(12, 0), inputRng.Offset(12, totalYear)).borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Range(inputRng.Offset(14, 0), inputRng.Offset(14, totalYear)).borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Range(inputRng.Offset(13, 0), inputRng.Offset(14, 0))
        .HorizontalAlignment = xlRight
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0.499984740745262
        .Font.Bold = True
    End With
    

    With Range(inputRng.Offset(1, 1), inputRng.Offset(14, totalYear))
        .Font.Name = "Arial"
        .Font.Size = 9
        .NumberFormat = "0.00%;-0.00%;"
    End With

   
    With Range(inputRng.Offset(1, 0), inputRng.Offset(13, 0)).Font
        .Bold = True
        '.ThemeColor = xlThemeColorDark1
        '.TintAndShade = 0.499984740745262
    End With
    
    With Range(inputRng.Offset(0, 1), inputRng.Offset(0, totalYear)).Font
        .Bold = True
         '.ThemeColor = xlThemeColorDark1
        '.TintAndShade = 0.499984740745262
    End With
    
    Range(inputRng, inputRng.Offset(15, totalYear)).Columns.AutoFit
    
    Unload Me
    Exit Sub
DateError:
    MsgBox "The date range does not look right!", vbCritical
    Exit Sub

End Sub



