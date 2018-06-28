VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSortColumns 
   Caption         =   "UserForm1"
   ClientHeight    =   7548
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9840
   OleObjectBlob   =   "ufSortColumns.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "ufSortColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private wbTarget As Workbook
Private wsTarget As Worksheet
Private HeaderRowNr As Integer
Private ChangesDone As Boolean
Dim cmdArray() As New cChkBoxPreset

' All Controls with Tag As Integer where bit 1 is set will be shown if ShowBackEnd is True
Private ShowBackEnd As Boolean
Private cbPresetList_Input_Given As Boolean



Private Sub btnClose_Click()
    Unload Me
End Sub


Private Sub btnHideAllColumns_Click()
    On Error GoTo ErrorHandler  ' in case the sheet is protected
    wsTarget.UsedRange.EntireColumn.Hidden = True
    
    On Error GoTo -1
    UpdateCheckBoxes
    Exit Sub
ErrorHandler:
    SysLog "btnHideAllColumns_Click", "Error"
End Sub


Private Sub btnLoadPresets_Click()
    LoadPresets
End Sub


Private Sub btnSaveAs_Click()
    SysLog "btnSaveAs_Click()"
    Dim Section As String
    Section = wbTarget.Name & "\" & wsTarget.Name
    
    SaveSetting AppName, Section, "PresetCurrent", tbSaveAs.text
    cbPresetList.AddItem tbSaveAs.text
    Dim i As Integer, sPresetList As String
    
    sPresetList = ""
    For i = 0 To cbPresetList.ListCount - 1
        sPresetList = sPresetList & cbPresetList.List(i) & ";"
    Next
    If AnyThing(sPresetList) Then
        sPresetList = Mid(sPresetList, 1, Len(sPresetList) - 1)
        SaveSetting AppName, Section, "PresetList", sPresetList
    Else
        SysLog "btnSaveAs_Click()", "sPresetList Nothing", sPresetList
    End If
    If cbPresetList.ListCount - 1 > -1 Then
        cbPresetList.ListIndex = cbPresetList.ListCount - 1
    End If
    SavePresets
    
    ChangesDone = False
    laPreset.Caption = "Preset navn (har gjort endringer: " & BoolToStr(ChangesDone) & ")"
End Sub


Private Sub btnShowAllColumns_Click()
    On Error GoTo ErrorHandler  ' in case the sheet is protected
    wsTarget.UsedRange.EntireColumn.Hidden = False
    
    On Error GoTo -1
    UpdateCheckBoxes
    Exit Sub
ErrorHandler:
    SysLog "btnHideAllColumns_Click", "Error"
End Sub


Private Sub btnUpdateCheckBoxes_Click()
    UpdateCheckBoxes
End Sub


Public Sub UserForm_Change()
    ChangesDone = True
    laPreset.Caption = "Preset navn (har gjort endringer: " & BoolToStr(ChangesDone) & ")"
End Sub


Private Sub cbPresetList_Change()
    
    If cbPresetList_Input_Given Then
        cbPresetList_Input_Given = False
        SysLog "cbPresetList_Change()", "cbPresetList_Input_Given"
        Dim Section, Key, Setting, HeaderName, order As String
        Section = wbTarget.Name & "\" & wsTarget.Name
        
        Dim ScrollColumn As Integer
        ScrollColumn = ActiveWindow.ScrollColumn
        SaveSetting AppName, Section, "PresetCurrent", cbPresetList.Value
        LoadPresets
        If ActiveWindow.ActiveSheet.Name = wsTarget.Name Then ActiveWindow.ScrollColumn = ScrollColumn
    End If
End Sub

Private Sub cbPresetList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    cbPresetList_Input_Given = True
End Sub

Private Sub cbPresetList_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cbPresetList_Input_Given = True
End Sub

Private Sub UserForm_Initialize()
    Project_Initialize
    SysLog "UserForm_Initialize()"
    Set wbTarget = Application.ActiveWorkbook
    Set wsTarget = Application.ActiveWorkbook.ActiveSheet
    
    If AnyThing(AppName) Then
        Me.StartUpPosition = 0  ' vbStartUpManual
        Me.Top = GetSetting(AppName:=AppName, Section:=Me.Name, Key:="Top", Default:="25")
        Me.Left = GetSetting(AppName:=AppName, Section:=Me.Name, Key:="Left", Default:="25")
    Else
        SysLog "UserForm_Initialize()", "AppName Not Anything"
    End If
    ChangesDone = False
    cbPresetList_Input_Given = False
    

    Dim item As Control, i As Integer
    ShowBackEnd = CBool(GetSetting(AppName, Me.Name, "ShowBackEnd", False))
    For Each item In Controls
        If IsNumeric(item.Tag) Then
            i = CInt(item.Tag)
            item.Visible = CBool(i And 1 And ShowBackEnd)
        End If
    Next
    
    If AnyThing(wsTarget) Then
        Me.Caption = wbTarget.Name & " - " & wsTarget.Name
    
        HeaderRowNr = FindSheetHeader(wsTarget)
        If AnyThing(HeaderRowNr) Then
            GenerateCheckBoxes
        Else
            SysLog "UserForm_Initialize()", "HeaderRowNr Not Anything"
        End If
    Else
        SysLog "UserForm_Initialize()", "wsTarget Not Anything"
    End If
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If ChangesDone Then
        Dim result As VbMsgBoxResult
        result = MsgBox("Ønsker du å lagre endret preset?", vbYesNoCancel, "Lagre Preset?")
        SysLog "UserForm_QueryClose()", "ChangesDone = " & CStr(ChangesDone), CInt(result)
        Select Case result
            Case vbOK      ' 1   OK
            Case vbCancel  ' 2   Cancel
                Cancel = True
            Case vbAbort   ' 3   Abort
            Case vbRetry   ' 4   Retry
            Case vbIgnore  ' 5   Ignore
            Case vbYes     ' 6   Yes
                Dim Section As String
                Section = wbTarget.Name & "\" & wsTarget.Name
                SaveSetting AppName, Section, "PresetCurrent", cbPresetList.Value
                SavePresets
            Case vbNo      ' 7   No
            Case Else
                SysLog "UserForm_QueryClose()", "ChangesDone = " & CStr(ChangesDone), "Case Else"
        End Select
    End If

    If AnyThing(AppName) Then
            SaveSetting AppName, Me.Name, "Top", Me.Top
            SaveSetting AppName, Me.Name, "Left", Me.Left
    Else
        SysLog "UserForm_QueryClose()", "AppName Not Anything"
    End If
End Sub



Private Sub GenerateCheckBoxes()
    If AnyThing(FindControlByName(Me.Controls, "chBoxGen0", "checkbox")) Then
        SysLog "GenerateCheckBoxes()", "already run, exiting sub"
        Exit Sub
    End If
    
    Dim MyRng, cell As Range, TargetFrame As Frame
    Set MyRng = wsTarget.UsedRange.rows(HeaderRowNr)
    Set TargetFrame = FrameFields
    
    Dim i, BoxPrRow, WidPrBox, HeiPrBox As Integer
    BoxPrRow = 4
    HeiPrBox = 18
    WidPrBox = TargetFrame.Width / BoxPrRow
    i = 0
    Dim chkBox As MSForms.CheckBox, Caption As String
    Dim Row, Col As Integer
    For Each cell In MyRng.Cells
        If AnyThing(cell) And AnyThing(cell.Value) Then
            Row = i \ BoxPrRow
            Col = i Mod BoxPrRow
            Set chkBox = TargetFrame.Controls.Add("Forms.CheckBox.1", "chBoxGen" & i, False)
            If AnyThing(chkBox) Then
                chkBox.Value = Not cell.EntireColumn.Hidden
                chkBox.WordWrap = False
                Caption = cell.Value
'                Caption = Replace(Caption, Chr(13) + Chr(10), " ")
'                Caption = Replace(Caption, Chr(13), " ")
'                Caption = Replace(Caption, Chr(10), " ")

                chkBox.Caption = Caption
                chkBox.Left = 8 + WidPrBox * Col
                chkBox.Top = 8 + HeiPrBox * Row
                chkBox.Visible = True
            End If
            ReDim Preserve cmdArray(0 To i)
            Set cmdArray(i).CheckBoxEvents = chkBox
            i = i + 1
        End If
    Next
    Set chkBox = Nothing
    
    LoadPresets
End Sub


Public Sub ShowHideColumnByName(Value As Boolean, HeaderName As String)
    If AnyThing(wsTarget) Then
        If AnyThing(HeaderRowNr) Then
            With wsTarget
                Dim isProtected As Boolean: isProtected = .ProtectContents
                Dim strSearch As String
                Dim aCell As Range
        
                If isProtected Then
                    .Unprotect
                End If
                ' The problem is in the .Find() call. Using LookIn:=xlValues won't find hidden cells.
                ' Change it to LookIn:=xlFormulas and it should work.
                Set aCell = .rows(HeaderRowNr).Find(What:=HeaderName, LookIn:=xlFormulas, _
                LookAt:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
        
                If AnyThing(aCell) Then
                    aCell.EntireColumn.Hidden = Not Value
                    ArrangeButtons
                Else
                    SysLog "ShowHideColumnByName()", "aCell Not Anything"
                End If
        
                If isProtected Then
                    .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
                End If
            End With
        Else
            SysLog "ShowHideColumnByName()", "HeaderRowNr Not Anything"
        End If
    Else
        SysLog "ShowHideColumnByName()", "wsTarget Not Anything"
    End If
End Sub


Public Sub UpdateCheckBoxes()
    If AnyThing(wsTarget) And AnyThing(HeaderRowNr) Then
        Dim MyRng, cell As Range
        Set MyRng = wsTarget.UsedRange.rows(HeaderRowNr)
        Dim chkBox As MSForms.CheckBox, item As Control

        For Each cell In MyRng.Cells
            Set chkBox = Nothing
            For Each item In Controls
                If LCase(TypeName(item)) = "checkbox" Then
                    If LCase(item.Caption) = LCase(cell.Value) Then
                        Set chkBox = item
                    End If
                End If
            Next item
            If AnyThing(chkBox) Then
                chkBox.Value = Not cell.EntireColumn.Hidden
            End If
        Next
    Else
        SysLog "UpdateCheckBoxes()", "AnyThing(wsTarget) And AnyThing(HeaderRowNr) Failed"
    End If
End Sub


Public Sub LoadPresets()
    SysLog "LoadPresets()"

    Dim Section, Key, Setting, HeaderName, order As String
    Dim Settings As Variant
    Dim i, splt, rdr As Integer
    Dim SkipPresetLoad, Value As Boolean
    
    Section = wbTarget.Name & "\" & wsTarget.Name
    SkipPresetLoad = GetSetting(AppName, Section, "SkipPresetLoad", False)
    If SkipPresetLoad Then
        SysLog "LoadPresets()", "SkipPresetLoad = " & CStr(SkipPresetLoad)
        Exit Sub
    End If
    
    Dim sPresetCurrent, sPresetList As String
    Dim aPresetList() As String
    
    sPresetCurrent = GetSetting(AppName, Section, "PresetCurrent", "")
    cbPresetList.Value = sPresetCurrent
    sPresetList = GetSetting(AppName, Section, "PresetList", "")
    aPresetList = Split(sPresetList, ";")
    cbPresetList.List = aPresetList
    
    
    Settings = GetAllSettings(AppName, Section & "\" & sPresetCurrent)
    If AnyThing(sPresetCurrent) And AnyThing(Settings) Then
        SysLog "LoadPresets()", "Settings[" & LBound(Settings) & ":" & UBound(Settings) & "] found at " & Section & "\" & sPresetCurrent
        For i = LBound(Settings, 1) To UBound(Settings, 1)
            Key = Settings(i, 0)
            splt = InStr(1, Key, " ")
            order = Mid(Key, 1, splt)
            HeaderName = Mid(Key, splt + 1)
            Setting = Settings(i, 1)

            If IsNumeric(order) And splt = 4 Then
                Value = Not CBool(Setting)
                ' CStr needs to be used in argument input to avoid ByRef argument type mismatch
                ' https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/byref-argument-type-mismatch
                 ShowHideColumnByName Value, CStr(HeaderName)
            End If
        Next
        UpdateCheckBoxes
    ElseIf Not AnyThing(sPresetCurrent) Then
        SysLog "LoadPresets()", "sPresetCurrent Not Anything"
    Else
        SysLog "LoadPresets()", "Section Not Anything"
    End If
End Sub

Public Sub SavePresets()
    SysLog "SavePresets()"
    Dim Section As String, SkipPresetSave As Boolean
    
    Section = wbTarget.Name & "\" & wsTarget.Name
    SkipPresetSave = GetSetting(AppName, Section, "SkipPresetSave", False)
    If SkipPresetSave Then
        SysLog "SavePresets()", "SkipPresetSave = " & CStr(SkipPresetSave)
        Exit Sub
    End If
    
    If AnyThing(cbPresetList.Value) Then
        Section = Section & "\" & cbPresetList.Value
    
        Dim MyRng, cell As Range
        Set MyRng = wsTarget.rows(HeaderRowNr)
        
        Dim i, z As Integer
        Dim Key, Setting As String
        
        i = 0
        Dim chkBox As MSForms.CheckBox, Caption As String
    '    Dim Row, Col As Integer
        For Each cell In MyRng.Cells
            If AnyThing(cell) And AnyThing(cell.Value) Then
                Key = Format(i, "000") & " " & cell.Value
                Setting = CStr(cell.EntireColumn.Hidden)
                SaveSetting AppName, Section, Key, Setting
                i = i + 1
            End If
        Next
    Else
        SysLog "SavePresets()", "cbPresetList.Value Nothing"
    End If
End Sub


Public Sub ArrangeButtons()
    If AnyThing(wsTarget) Then
        With wsTarget
            Dim objX As Object, i, btnWid, btnWidSpc As Integer
            i = 0
            btnWid = 96
            btnWidSpc = 12  ' Button Width Spaceing
            
            For Each objX In .OLEObjects
                If TypeName(objX.Object) = "CommandButton" Then
                    If objX.Visible Then
                        objX.Left = btnWidSpc + i * (btnWid + btnWidSpc)
                        i = i + 1
                    End If
                End If
            Next
        End With
    End If
End Sub
