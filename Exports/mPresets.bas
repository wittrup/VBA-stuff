Attribute VB_Name = "mPresets"
Option Explicit


Public Sub ApplyPreset(Preset As String)
    Dim shtPresets As Worksheet
    Dim rngHeader, rngPresets, cell As Range
    Set shtPresets = Sheet02
    Set rngPresets = shtPresets.Range("A:A").Find(What:=Preset)
    Set rngHeader = shtPresets.UsedRange.Rows(1)
    
    If AnyThing(rngPresets) Then
        Application.ScreenUpdating = False

        Dim colStart As Integer: colStart = 0

        For Each cell In rngHeader.Cells
            If AnyThing(cell) And AnyThing(cell.Value) And colStart > 0 Then
                ShowHideColumnByName CBool(shtPresets.Cells(rngPresets.Row, cell.Column).Text <> vbNullString), CStr(cell.Value)
            End If
            colStart = colStart + 1
        Next
        Sheet01.ArrangeButtons
    Else
        Sheet01.ShowAllColumns
        Debug.Print Now, "ApplyPreset() FoundRow Nothing", Preset
    End If
End Sub


Public Sub ShowHideColumnByName(Value As Boolean, Caption As String)
    Dim Row As Integer
    Row = 2
    
    With Sheet01
        Dim isProtected As Boolean: isProtected = .ProtectContents
        Dim strSearch As String
        Dim aCell As Range
        
        If isProtected Then
            .Unprotect
        End If
        ' The problem is in the .Find() call. Using LookIn:=xlValues won't find hidden cells.
        ' Change it to LookIn:=xlFormulas and it should work.
        Set aCell = .Rows(Row).Find(What:=Caption, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    
        If AnyThing(aCell) Then
            aCell.EntireColumn.Hidden = Not Value
        End If
        
        If isProtected Then
            .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        End If
    End With
End Sub

