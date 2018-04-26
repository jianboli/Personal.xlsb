Attribute VB_Name = "MergeSelectedRange"
Option Explicit

'   Stores info about current selection
Type SaveRange
    Val As Variant
    Addr As String
    hAlignment As Integer
    vAlignment As Integer
    leftBorder As Integer
    rightBorder As Integer
    topBorder As Integer
    bottomBorder As Integer
End Type
Dim OldWorkbook As Workbook
Dim OldSheet As Worksheet
Dim OldSelection() As SaveRange



Sub MergeSelectedRange()
    Dim rng As Range
    Dim sh As Worksheet
    Set rng = Selection
    Set sh = Application.ActiveSheet
    
    If rng.Columns.Count = 0 Or rng.Rows.Count <= 1 Then
        MsgBox "Please select something before click this button!", vbOKOnly + vbExclamation, "Dr. Bob"
        Exit Sub
    End If
    
    'Record current status
    ReDim OldSelection(rng.Count)
    Set OldWorkbook = ActiveWorkbook
    Set OldSheet = ActiveSheet
    Dim i As Integer, cell
    i = 0
    For Each cell In Selection
        i = i + 1
        OldSelection(i).Addr = cell.Address
        OldSelection(i).Val = cell.Formula
        OldSelection(i).hAlignment = cell.HorizontalAlignment
        OldSelection(i).vAlignment = cell.VerticalAlignment
        OldSelection(i).leftBorder = cell.borders(xlEdgeLeft).LineStyle
        OldSelection(i).rightBorder = cell.borders(xlEdgeRight).LineStyle
        OldSelection(i).topBorder = cell.borders(xlEdgeTop).LineStyle
        OldSelection(i).bottomBorder = cell.borders(xlEdgeBottom).LineStyle
    Next cell
    
    
    Dim j As Integer, startRow As Integer
    Dim rngToBeMerged As Range
    Dim content
    Application.DisplayAlerts = False
    For i = 1 To rng.Columns.Count
        content = rng.Cells(1, i).Text
        startRow = 1
        For j = 2 To rng.Rows.Count
            If (rng.Cells(j, i).Text <> content) Then
                Set rngToBeMerged = sh.Range(rng.Cells(startRow, i), rng.Cells(j - 1, i))
                rngToBeMerged.Merge
                rngToBeMerged.HorizontalAlignment = xlCenter
                rngToBeMerged.VerticalAlignment = xlCenter
                rngToBeMerged.BorderAround Weight:=xlMedium
                startRow = j
                content = rng.Cells(j, i).Text
            End If
        Next j
        Set rngToBeMerged = sh.Range(rng.Cells(startRow, i), rng.Cells(j - 1, i))
        rngToBeMerged.Merge
        rngToBeMerged.HorizontalAlignment = xlCenter
        rngToBeMerged.VerticalAlignment = xlCenter
        rngToBeMerged.BorderAround Weight:=xlMedium
    Next i
    Application.DisplayAlerts = True

'   Specify the Undo Sub
    Application.OnUndo "Undo the Merge", "UndoMergeSelectedRange"
End Sub

Sub UndoMergeSelectedRange()
'   Undoes the effect of the ZeroRange sub
    
'   Tell user if a problem occurs
    On Error GoTo Problem

'   Make sure the correct workbook and sheet are active
    OldWorkbook.Activate
    OldSheet.Activate
    Dim rMerged As Range
    Dim v
'   Restore the saved information
    Dim i As Integer
    For i = 1 To UBound(OldSelection)
        If Range(OldSelection(i).Addr).MergeCells Then
            Set rMerged = Range(OldSelection(i).Addr).MergeArea
            rMerged.MergeCells = False
        End If
    Next i
    
    Dim cell As Range
    For i = 1 To UBound(OldSelection)
        Set cell = Range(OldSelection(i).Addr)
        cell.Formula = OldSelection(i).Val
        cell.HorizontalAlignment = OldSelection(i).hAlignment
        cell.VerticalAlignment = OldSelection(i).vAlignment
        cell.borders(xlEdgeLeft).LineStyle = OldSelection(i).leftBorder
        cell.borders(xlEdgeRight).LineStyle = OldSelection(i).rightBorder
        cell.borders(xlEdgeTop).LineStyle = OldSelection(i).topBorder
        cell.borders(xlEdgeBottom).LineStyle = OldSelection(i).bottomBorder
    Next i
    Exit Sub
'   Error handler
Problem:
    MsgBox "Can't undo"
End Sub




