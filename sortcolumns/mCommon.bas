Attribute VB_Name = "mCommon"

Option Explicit

Public UserName, ComputerName, AppName As String
Private Const Verbose = False


Public Sub Project_Initialize()
    If Not AnyThing(UserName) Then
        UserName = apicGetUserName
        SysLog "UserName     set", UserName
    End If
    If Not AnyThing(ComputerName) Then
        ComputerName = apicGetComputerName
        SysLog "ComputerName set", ComputerName
    End If
    If Not AnyThing(AppName) Then
        AppName = UCase(Mid(ThisWorkbook.Name, 1, mCommon.Min(InStrRev(ThisWorkbook.Name, "."), Len(ThisWorkbook.Name)) - 1))
        SysLog "AppName      set", AppName
    End If
End Sub


Public Sub SysLog(ParamArray var() As Variant)
    Debug.Print Now, UserName & "@" & ComputerName, AppName, Join(var, vbTab)
End Sub


Public Function FindSheetHeader(sh As Worksheet) As Integer
    Dim s As String
    If sh.AutoFilterMode Then  ' If AutoFilter applied, assume top of range is sheet header row
        s = sh.AutoFilter.Range.Address
        FindSheetHeader = onlyDigits(Mid(s, 1, InStr(s, ":")))
    End If
End Function


Public Function AnyThing(Value As Variant) As Boolean
' This function check is the value contains anything else then null.
' For strings space also counts as an empty string "Trim()"
    Dim result As Boolean, typNam As String
    typNam = TypeName(Value)
    If Verbose Then Debug.Print typNam, IsArray(Value)
    If IsArray(Value) Then
        result = CBool(ArrayLen(Value))
    Else
        Select Case typNam
        Case "String"
            result = CBool(Trim(Value & vbNullString) <> vbNullString)
        Case "Empty"
            result = False
        Case "Integer"
            result = CBool(Value)
        Case Else
            result = Not Value Is Nothing
        End Select
    End If
    AnyThing = result
End Function


Public Function onlyDigits(s As String) As String
    ' https://stackoverflow.com/questions/7239328/how-to-find-numbers-from-a-string
    Dim retval As String, i As Integer
    retval = ""
    
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next
    onlyDigits = retval
End Function


Public Function Min(i1, i2 As Integer) As Integer
    If i1 > i2 Then
        Min = i2
    Else
        Min = i1
    End If
End Function


Public Function FindControlByName(Controls As Object, ControlName, TypeStr As String) As Control
    Dim item As Control
    ControlName = LCase(ControlName)
    TypeStr = LCase(TypeStr)
    Set FindControlByName = Nothing
    
    For Each item In Controls
'           TypeOf item is CheckBox - not working?
        If LCase(TypeName(item)) = TypeStr And LCase(item.Name) = ControlName Then
            Set FindControlByName = item
        End If
    Next item
End Function


Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

