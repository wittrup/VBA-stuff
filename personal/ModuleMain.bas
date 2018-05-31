Attribute VB_Name = "ModuleMain"

Option Explicit


Sub CLEAR_ALL_FILTERS_ACTIVE_SHEET()
    If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
End Sub


Sub DELETE_EMPTY_ROWS()
  Dim r As Range, rows As Long, i As Long
  Set r = ActiveSheet.UsedRange
  rows = r.rows.Count
  For i = rows To 1 Step (-1)
    If WorksheetFunction.CountA(r.rows(i)) = 0 Then r.rows(i).Delete
  Next
End Sub


Sub AUTOFIT_COLUMNS()
    ActiveSheet.Cells.EntireColumn.AutoFit
End Sub


Sub INC_COLS_SIZE()
    Dim sht As Worksheet, cell As Range
    Set sht = ActiveSheet
    
    For Each cell In sht.UsedRange.Columns
        cell.ColumnWidth = cell.ColumnWidth * 1.15
    Next cell
End Sub


Sub COPY_SHEET_TO_NEW()
    ActiveSheet.Copy
End Sub
