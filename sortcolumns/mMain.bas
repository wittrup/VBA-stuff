Attribute VB_Name = "mMain"

Option Explicit

Dim arrUfSortColumns() As New ufSortColumns


Sub ufSortColumns_Show()
'    Load ufSortColumns
    ufSortColumns.Show
End Sub


Sub LoadPreset()
    ufSortColumns.LoadPresets
End Sub


Sub SavePreset()
    ufSortColumns.SavePresets
End Sub
