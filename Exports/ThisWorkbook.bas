VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Workbook_Open()
    Debug.Print Now, "Workbook_Open()", apicGetUserName + "@" + apicGetComputerName
    ApplyPreset GetSetting(AppName, "\", "Preset", "ALL")
    Sheet01.ArrangeButtons
    Application.ScreenUpdating = True
End Sub
