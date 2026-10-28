Attribute VB_Name = "PlannerMacros"
' Optional buttons for the release workbooks (B-704). Import this file into a copy of the planner saved as .xlsm
' (Developer > Visual Basic > File > Import File), then add buttons or Quick Access Toolbar entries for the macros.
' Every macro only automates something the .xlsx already does by hand (TakeSnapshot sheet, Start > Inserting
' tasks); the workbook never needs it. Everything is found by defined name or table, never by cell address, so the
' macros keep working when rows and columns move.
Option Explicit

' TakeSnapshot > BASELINE: copy Tasks' BL Source columns and paste them as values onto the Baseline columns.
Public Sub SetBaseline()
    Dim source As Range, target As Range
    Set source = ThisWorkbook.Names("tsk_BLSource").RefersToRange
    Set target = ThisWorkbook.Names("tsk_Baseline").RefersToRange
    If MsgBox("Copy the current schedule into the Baseline columns (" & _
              IIf(ThisWorkbook.Names("TakeBaseline").RefersToRange.Value = "Yes", "every task", _
                  "tasks without a baseline") & ")?", vbOKCancel + vbQuestion, "Set baseline") <> vbOK Then Exit Sub
    target.Value = source.Value
    ThisWorkbook.Names("TakeBaseline").RefersToRange.Value = "No"
End Sub

' Template 1.2 (B-201): keep the current baseline's dates as Baseline 2 or 3 before re-baselining.
Public Sub KeepBaseline()
    Dim which As String
    which = InputBox("Keep the current baseline as Baseline 2 or 3?", "Keep a baseline", "2")
    If which <> "2" And which <> "3" Then Exit Sub
    ThisWorkbook.Names("tsk_Baseline" & which).RefersToRange.Value = _
        ThisWorkbook.Names("tsk_BaseDates").RefersToRange.Value
End Sub

' TakeSnapshot > SNAPSHOT: paste this status date's numbers as values under the last row of History.
Public Sub RecordSnapshot()
    Dim history As ListObject, row As Range, i As Long
    Set history = ThisWorkbook.Worksheets("History").ListObjects("StatusHistory")
    For i = 1 To history.ListRows.Count
        If IsEmpty(history.DataBodyRange.Cells(i, 1).Value) Then Exit For
    Next i
    If i > history.ListRows.Count Then history.ListRows.Add
    Set row = history.DataBodyRange.Rows(i)
    row.Resize(1, ThisWorkbook.Names("snap_Row").RefersToRange.Columns.Count).Value = _
        ThisWorkbook.Names("snap_Row").RefersToRange.Value
End Sub

' Start > Inserting tasks: a new row below the selected task, at its level, with the next free UID.
Public Sub InsertTask()
    Dim tasks As ListObject, at As Long, added As ListRow
    Set tasks = ThisWorkbook.Worksheets("Tasks").ListObjects("Tasks")
    at = SelectedTaskRow(tasks)
    If at = 0 Then Exit Sub
    Set added = tasks.ListRows.Add(at + 1)
    added.Range.Cells(1, tasks.ListColumns("UID").Index).Value = ThisWorkbook.Names("NextFreeUID").RefersToRange.Value
    added.Range.Cells(1, tasks.ListColumns("Level").Index).Value = _
        tasks.DataBodyRange.Cells(at, tasks.ListColumns("Level").Index).Value
    added.Range.Cells(1, tasks.ListColumns("Task Name").Index).Select
End Sub

' One level deeper / shallower for the selected tasks (Level is an input; the outline follows).
Public Sub IndentTask()
    ShiftLevel 1
End Sub

Public Sub OutdentTask()
    ShiftLevel -1
End Sub

' Show every task, or only the phases (Excel's outline buttons 1 / 9 on the Tasks rows).
Public Sub ExpandAll()
    ThisWorkbook.Worksheets("Tasks").Outline.ShowLevels RowLevels:=8
End Sub

Public Sub CollapseAll()
    ThisWorkbook.Worksheets("Tasks").Outline.ShowLevels RowLevels:=1
End Sub

Private Sub ShiftLevel(by As Long)
    Dim tasks As ListObject, cell As Range, levels As Range, r As Long
    Set tasks = ThisWorkbook.Worksheets("Tasks").ListObjects("Tasks")
    Set levels = tasks.ListColumns("Level").DataBodyRange
    For Each cell In Selection.Rows
        If Not Intersect(cell.EntireRow, tasks.DataBodyRange) Is Nothing Then
            r = cell.Row - tasks.HeaderRowRange.Row
            levels.Cells(r, 1).Value = Application.Max(1, Val(levels.Cells(r, 1).Value) + by)
        End If
    Next cell
End Sub

Private Function SelectedTaskRow(tasks As ListObject) As Long
    If Intersect(ActiveCell, tasks.DataBodyRange) Is Nothing Then
        MsgBox "Select a task on the Tasks sheet first.", vbInformation, "Insert task"
        Exit Function
    End If
    SelectedTaskRow = ActiveCell.Row - tasks.HeaderRowRange.Row
End Function
