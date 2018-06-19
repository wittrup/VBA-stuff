VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cChkBoxPreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public WithEvents CheckBoxEvents As MSForms.CheckBox
Attribute CheckBoxEvents.VB_VarHelpID = -1

Private Input_Given As Boolean


Private Sub Class_Initialize()
    Input_Given = False
End Sub


Private Sub CheckBoxEvents_Change()
'   Event order:    MouseDown   -> MouseUp  -> Change   -> Click
'                   KeyDown     -> KeyUp    -> Change   -> Click
    If Input_Given Then
        Input_Given = False
        ufSortColumns.ShowHideColumnByName CheckBoxEvents.Value, CheckBoxEvents.Caption
    End If
End Sub


Private Sub CheckBoxEvents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Input_Given = True
End Sub


Private Sub CheckBoxEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Input_Given = True
End Sub
