Attribute VB_Name = "mDeveloper"
Option Explicit



Sub Export()
    Dim objMyProj As VBProject, objVBComp As VBComponent, t As Long
    ' You need to add a reference to the VBA extensibility library.
    ' Click on Tools-References in the VBE, and scroll down and tick
    ' the entry for Microsoft Visual Basic for Applications Extensibility 5.3.
    
    ' Klarer tilgang til VBA-prosjektobjektmodellen
    ' https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/type-property-vba-add-in-object-model
    Set objMyProj = Application.VBE.ActiveVBProject
    
    ' Constant                  Value   Description
    ' vbext_ct_StdModule          1     Standard module
    ' vbext_ct_ClassModule        2     Class module
    ' vbext_ct_MSForm             3     Microsoft Form
    ' vbext_ct_ActiveXDesigner   11     ActiveX Designer
    ' vbext_ct_Document         100     Document Module

    For Each objVBComp In objMyProj.VBComponents
        t = objVBComp.Type
        If Not objVBComp.Name Like "X_*" Then
            objVBComp.Export ThisWorkbook.Path & "\Exports\" & objVBComp.Name & ".bas"
        End If
    Next
End Sub
