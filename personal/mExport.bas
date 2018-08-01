Attribute VB_Name = "mExport"

Option Explicit


Private Sub Export()
    Dim objMyProj As VBProject, objVBComp As VBComponent, Folder As String
    Dim t As Long, fileName As String, tempFile As String
    Dim fso, theFile
    ' https://stackoverflow.com/a/14994750/2029846
    
    ' You need to add a reference to the VBA extensibility library.
    ' Click on Tools-References in the VBE, and scroll down and tick
    ' the entry for Microsoft Visual Basic for Applications Extensibility 5.3.
    
    ' Alternativer -> Klareringssenter -> Makroinnstillinger -> Klarer tilgang til VBA-prosjektobjektmodellen
    ' https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/type-property-vba-add-in-object-model
    Set objMyProj = Application.VBE.ActiveVBProject
    
    ' Constant                  Value   Description
    ' vbext_ct_StdModule          1     Standard module
    ' vbext_ct_ClassModule        2     Class module
    ' vbext_ct_MSForm             3     Microsoft Form
    ' vbext_ct_ActiveXDesigner   11     ActiveX Designer
    ' vbext_ct_Document         100     Document Module
    
    Folder = LCase(Mid(ThisWorkbook.Name, 1, Min(Array(InStrRev(ThisWorkbook.Name, "."), Len(ThisWorkbook.Name))) - 1))
    Folder = ThisWorkbook.Path & "\" & Folder & "\"
    Debug.Print Folder
    createNewDirectory Folder
    
    For Each objVBComp In objMyProj.VBComponents
        t = objVBComp.Type
        fileName = Folder & objVBComp.Name & ".bas"
        tempFile = Environ("temp") & "\" & RemFilExt(ThisWorkbook.Name) & " - " & objVBComp.Name & ".bas"
        objVBComp.Export tempFile
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set theFile = fso.OpenTextFile(tempFile, 8, True)
        
        If t = vbext_ct_Document And theFile.Line = 10 Then  ' WorkSheets
        ElseIf t = vbext_ct_StdModule And theFile.Line = 2 Then  ' Modules
        ElseIf t = vbext_ct_MSForm And theFile.Line = 16 Then  ' UserForms
        Else
            ' TODO create checksum between files, if not match, copy from temp to file
            If FileExists(fileName) Then
                If mHash.FileToSHA256(tempFile) <> mHash.FileToSHA256(fileName) Then
                    Call fso.CopyFile(tempFile, fileName)
                End If
            Else
                Call fso.CopyFile(tempFile, fileName)
            End If
        End If
        Set fso = Nothing
    Next
    Set theFile = Nothing
End Sub


Private Sub createNewDirectory(directoryName As String)
    If Not DirExists(directoryName) Then
        MkDir (directoryName)
    End If
End Sub


Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
End Function


Public Function Min(sArray) As Integer
    Dim element As Variant
    Min = sArray(0)
 
    For Each element In sArray
        If Min > element Then
            Min = element
        End If
    Next element
End Function


Function RemFilExt(text As String) As String
    RemFilExt = Left(text, (InStrRev(text, ".", -1, vbTextCompare) - 1))
End Function


Function FileExists(ByVal FilePath As String) As Boolean
    Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function
