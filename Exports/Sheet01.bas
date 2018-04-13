VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub btnAddItem_Click()
    ufAddItem.Show
End Sub


Private Sub btnShowAllCol_Click()
    ShowAllColumns
    SaveSetting AppName, "\", "Preset", "ALL"
End Sub

Sub ShowAllColumns()
    Application.ScreenUpdating = False
    With Sheet01
        Dim isProtected As Boolean: isProtected = .ProtectContents
        .Unprotect
        
        .Columns.Hidden = False
        ArrangeButtons
        
        If isProtected Then
            .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        End If
    End With
    Application.ScreenUpdating = True
End Sub


Private Sub btnSortCol_Click()
    ufSort.Show
End Sub


Sub ArrangeButtons()
    With Sheet01
        Dim isProtected As Boolean: isProtected = .ProtectContents
        If isProtected Then
            .Unprotect
        End If
        
        Dim objX As Object, i, btnWid, btnWidSpc As Integer
        i = 0
        btnWid = 96
        btnWidSpc = 12  ' Button Width Spaceing
        
        For Each objX In .OLEObjects
            If typename(objX.Object) = "CommandButton" Then
                If objX.Visible Then
                    objX.Left = btnWidSpc + i * (btnWid + btnWidSpc)
                    i = i + 1
                End If
            End If
        Next

        If isProtected Then
            .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        End If
    End With
End Sub







