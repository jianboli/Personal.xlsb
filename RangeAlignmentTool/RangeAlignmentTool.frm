VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RangeAlignmentTool 
   Caption         =   "Range Alignment"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5250
   OleObjectBlob   =   "RangeAlignmentTool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RangeAlignmentTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type SaveRange
    Val As Variant
    Addr As String
End Type
Dim OldWorkbook As Workbook
Dim OldSheet As Worksheet
Dim OldSelection() As SaveRange
Dim FinalRange1 As Range
Dim FinalRange2 As Range

Private Sub btnCancel_Click()
    UndoAlignRange
End Sub

Private Sub btnOK_Click()
    Dim rng1 As Range, rng2 As Range
    Set rng1 = Range(RefEdit1.Value)
    Set rng2 = Range(RefEdit2.Value)
    
    'Record current status
    ReDim OldSelection(rng1.Count + rng2.Count)
    Set OldWorkbook = ActiveWorkbook
    Set OldSheet = ActiveSheet
    Dim i As Integer, cell
    i = 0
    For Each cell In rng1
        i = i + 1
        OldSelection(i).Addr = cell.Address
        OldSelection(i).Val = cell.Formula

    Next cell
    For Each cell In rng2
        i = i + 1
        OldSelection(i).Addr = cell.Address
        OldSelection(i).Val = cell.Formula
    Next cell
    
    
    If HeaderCheck.Value Then
        Set rng1 = rng1.Offset(1, 0).Resize(rng1.Rows.Count - 1)
        Set rng2 = rng2.Offset(1, 0).Resize(rng2.Rows.Count - 1)
    End If
    
    Dim col1 As Integer, col2 As Integer
    col1 = CInt(Column1.Text)
    col2 = CInt(Column2.Text)
    
    Dim Col1C As Range, Col2C As Range
    Set Col1C = rng1(, col1)
    Set Col2C = rng2(, col2)
    
    ' sort them based on the given column
    rng1.sort Col1C
    rng2.sort Col2C
    
    ' align them based on the given column
    'Dim i As Integer
    i = 1
    Do While (i <= rng1.Rows.Count And i <= rng2.Rows.Count)
        If (IsEmpty(rng1(i, col1)) And IsEmpty(rng2(i, col2))) Then
            Exit Sub
        End If
        Dim res
        Select Case (StrComp(rng1(i, col1), rng2(i, col2), vbTextCompare))
            Case -1
                rng2.Rows(i).Insert Shift:=xlDown
            Case 1
                rng1.Rows(i).Insert Shift:=xlDown
        End Select
        i = i + 1
    DoEvents
    Loop
    
    Set FinalRange1 = rng1
    Set FinalRange2 = rng2
    '   Specify the Undo Sub
    Application.OnUndo "Undo the Alignment", "UndoAlignRange"
End Sub


Private Sub UserForm_Initialize()
    Column1.Text = 1
    Column2.Text = 1
    HeaderCheck.Value = True
    
    RefEdit1.Value = Sheet1.Cells(1, 1)
    RefEdit2.Value = Sheet1.Cells(2, 1)
    
    
    Dim rng As Range
    If TypeName(Application.Selection) = "Range" Then
        Set rng = Application.Selection
        If rng.Count > 1 Then
            RefEdit1.Value = rng.Address
        End If
    End If
    
End Sub

Sub sort()
    Dim sh1 As Worksheet, sh2 As Worksheet
    Set sh1 = Sheet1
    Set sh2 = Sheet2
    Dim wid
    wid = Application.WorksheetFunction.CountA(sh1.Range("2:2"))
    Dim colnum As Integer
    colnum = 2
    Dim i As Integer
    i = 2
    
End Sub



Sub UndoAlignRange()
'   Undoes the effect of the ZeroRange sub
    
'   Tell user if a problem occurs
    On Error GoTo Problem
    
'   Make sure the correct workbook and sheet are active
    OldWorkbook.Activate
    OldSheet.Activate

    FinalRange1.Clear
    FinalRange2.Clear
    
    Dim cell As Range
    Dim i As Integer
    For i = 1 To UBound(OldSelection)
        Set cell = Range(OldSelection(i).Addr)
        cell.Formula = OldSelection(i).Val
    Next i
    Exit Sub
'   Error handler
Problem:
    MsgBox "Can't undo"
End Sub

Private Sub UserForm_Terminate()
    ' remember my previous selection
    If Not IsEmpty(RefEdit1.Text) Then
        Sheet1.Cells(1, 1) = RefEdit1.Value
    End If
    If Not IsEmpty(RefEdit2.Text) Then
        Sheet1.Cells(2, 1) = RefEdit2.Value
    End If
End Sub
