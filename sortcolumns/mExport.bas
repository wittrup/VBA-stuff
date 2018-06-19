Attribute VB_Name = "mExport"

Option Explicit


Private Sub Export()
    Dim objMyProj As VBProject, objVBComp As VBComponent, Folder As String
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
    Folder = ThisWorkbook.path & "\" & Folder & "\"
    createNewDirectory Folder
    Debug.Print Folder
    
    For Each objVBComp In objMyProj.VBComponents
        objVBComp.Export Folder & objVBComp.Name & ".bas"
    Next
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
