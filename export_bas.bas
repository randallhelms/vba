Attribute VB_Name = "Module118"
Sub export_bas()
     
     ' reference to extensibility library
     
    Dim objMyProj As VBProject
    Dim objVBComp As VBComponent
     
    Set objMyProj = Application.VBE.ActiveVBProject
     
    For Each objVBComp In objMyProj.VBComponents
        If objVBComp.Type = vbext_ct_StdModule Then
            objVBComp.Export "C:\Users\Randall\Google Drive\data\vba\" & objVBComp.Name & ".bas"
        End If
    Next
     
End Sub
