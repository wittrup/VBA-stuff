VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSort 
   Caption         =   "UserForm1"
   ClientHeight    =   9528
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11916
   OleObjectBlob   =   "ufSort.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "ufSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnClose_Click()
' CommandButton Close
    Unload Me
End Sub


Private Sub btnSelect_Click()
' CommandButton Select / De-Select All

End Sub


Private Sub cbPresetSelect_Change()
' ComboBox cbPresetSelect
    SaveSetting AppName, "\", "Preset", cbPresetSelect.Value
    Dim ws As Worksheet, FoundRow, FoundCol As Range, item As Control, chkBox As MSForms.CheckBox
    Set ws = Sheet02
    Set FoundRow = ws.Range("A:A").Find(What:=cbPresetSelect.Value)
    
    If AnyThing(FoundRow) Then
        Application.ScreenUpdating = False
        For Each item In Me.Controls
'           TypeOf item is CheckBox - not working?
            If typename(item) = "CheckBox" Then
                Set chkBox = item
                item.Value = False
                Set FoundCol = ws.UsedRange.Find(item.Caption, , xlFormulas, xlWhole, , , False)
                If AnyThing(FoundCol) Then
                    chkBox.Value = CBool(ws.Cells(FoundRow.Row, FoundCol.Column).Text <> vbNullString)
                    ShowHideColumnByName chkBox.Value, chkBox.Caption
                End If
            End If
        Next
        Sheet01.ArrangeButtons
    Else
        Debug.Print Now, "FoundRow Nothing", cbPresetSelect.Value
    End If
End Sub



Private Sub UserForm_Initialize()
    Dim SettingLeft, SettingTop As Integer
    SettingTop = GetSetting(AppName:=AppName, section:=Name, Key:="Top", Default:="25")
    SettingLeft = GetSetting(AppName:=AppName, section:=Name, Key:="Left", Default:="25")
    
    With Me
        Top = SettingTop
        Left = SettingLeft
    End With
    
    GenerateCheckBoxes
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    With Me
        SaveSetting AppName, Name, "Top", Top
        SaveSetting AppName, Name, "Left", Left
        If CloseMode = vbFormControlMenu Then Cancel = True
        Hide
    End With
End Sub


Private Sub GenerateCheckBoxes()
    Dim MySht As Worksheet
    Dim MyRng, cell As Range

    If AnyThing(FindControlByName(Me.Controls, "chBoxGen0", "checkbox")) Then
        Debug.Print Now, "GenerateCheckBoxes may only run once"
        Exit Sub
    End If

    Set MySht = Sheet02
    Set MyRng = MySht.UsedRange.Rows(1)
    Dim TargetFrame As Frame
    Set TargetFrame = FrameFields
    
    Dim i, wid, BoxPrRow, WidPrBox, HeiPrBox, z As Integer
    wid = TargetFrame.Width
    BoxPrRow = 4
    HeiPrBox = 18
    WidPrBox = wid / BoxPrRow
    i = 0
    z = 0
    Dim chkBox As MSForms.CheckBox, Caption As String
    Dim Row, Col As Integer
    For Each cell In MyRng.Cells
        If AnyThing(cell) And AnyThing(cell.Value) And z > 0 Then
            Row = i \ BoxPrRow
            Col = i Mod BoxPrRow
            Set chkBox = TargetFrame.Controls.Add("Forms.CheckBox.1", "chBoxGen" & i)
            If AnyThing(chkBox) Then
                chkBox.WordWrap = False
                Caption = cell.Value
'                Caption = Replace(Caption, Chr(13) + Chr(10), " ")
'                Caption = Replace(Caption, Chr(13), " ")
'                Caption = Replace(Caption, Chr(10), " ")

                chkBox.Caption = Caption
                chkBox.Left = 8 + WidPrBox * Col
                chkBox.Top = 8 + HeiPrBox * Row
            End If
            i = i + 1
        End If
        z = z + 1
    Next
End Sub




