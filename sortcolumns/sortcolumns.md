##### Usage:

1. Make sure ```SORTCOLUMNS.XLSB``` is place in ```%AppData%\Microsoft\Excel\XLSTART```

2. Create a excel spreadsheet with autofilter.

3. Add an Active-X button for opening form (in this case ```btnSortCol```), use this code for button click:
```Visual Basic
Private Sub btnSortCol_Click()
    Application.Run "SORTCOLUMNS.XLSB!ufSortColumns_Show"
End Sub
```

4. Add load and saving code to WorkBook code:
```Visual Basic
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.Run "SORTCOLUMNS.XLSB!SavePreset"
End Sub

Private Sub Workbook_Open()
    Application.Run "SORTCOLUMNS.XLSB!LoadPreset"
End Sub
```