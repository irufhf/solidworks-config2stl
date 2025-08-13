Attribute VB_Name = "modConfig2stl"
' ========================================================================
' modConfig2stl.bas
'
' Description:
'   Exports all configurations from the active SolidWorks part or assembly
'   as individual STL files.
'
' Features:
'   - STL filenames are based on configuration names (invalid characters replaced)
'   - Output folder is created as a subdirectory ("STL_Exports") in the model's location
'   - Uses the last-used STL export settings (resolution, units, binary/ascii)
'   - Supports both parts and assemblies (assemblies exported as a single STL)
'
' Notes:
'   - Ensure the document is saved before running (to determine the output path)
'   - Modify OUTPUT_SUBFOLDER constant if a different export folder name is desired
'   - STL export options must be set manually prior to running the macro
'
' Author: Daan Verhoeff
' Date:   13-08-2025
' ========================================================================

Option Explicit

Const OUTPUT_SUBFOLDER As String = "STL_Exports"
Const swExportSTL As Long = 14 ' Constant for STL export format

Dim swApp As SldWorks.SldWorks

Sub main()

    ' Initialize SolidWorks application object
    Set swApp = Application.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    ' Validate that a document is open
    If swModel Is Nothing Then
        MsgBox "Please open a part or assembly before running this macro.", vbExclamation
        Exit Sub
    End If
    
    ' Ensure the document has been saved (required for determining output folder)
    Dim docPath As String
    docPath = swModel.GetPathName
    If Len(docPath) = 0 Then
        MsgBox "Please save the document before exporting.", vbExclamation
        Exit Sub
    End If
    
    ' Build the output folder path (create subfolder under model location)
    Dim modelFolder As String
    modelFolder = Left$(docPath, InStrRev(docPath, "\"))
    
    Dim outFolder As String
    outFolder = modelFolder & OUTPUT_SUBFOLDER & "\"
    CreateFolderIfNeeded outFolder
    
    ' Get all configuration names
    Dim vConfs As Variant
    vConfs = swModel.GetConfigurationNames
    If IsEmpty(vConfs) Then
        MsgBox "No configurations found in this document.", vbExclamation
        Exit Sub
    End If
    
    Dim i As Long, n As Long
    n = UBound(vConfs) - LBound(vConfs) + 1
    
    Dim swExt As SldWorks.ModelDocExtension
    Set swExt = swModel.Extension
    
    ' Attempt to get STL export data object (allows control over STL export options)
    Dim exportData As Object
    On Error Resume Next
    Set exportData = swApp.GetExportFileData(swExportSTL)
    On Error GoTo 0
    
    Dim errNum As Long, warnNum As Long
    Dim savedCount As Long
    
    ' Assemblies: export as one combined STL file
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSTLComponentsIntoOneFile, True
    
    ' Loop through all configurations
    For i = LBound(vConfs) To UBound(vConfs)
        Dim confName As String
        confName = CStr(vConfs(i))
        
        ' Activate the configuration and rebuild the model
        swModel.ShowConfiguration confName
        swModel.ForceRebuild3 False
        
        ' Sanitize configuration name for use as a filename
        Dim safeName As String
        safeName = SanitizeFileName(confName)
        
        ' Build the output file path
        Dim outFile As String
        outFile = outFolder & safeName & ".stl"
        
        ' Export the STL file (silent mode)
        If Not exportData Is Nothing Then
            ' Using ExportFileData object (preferred method)
            swExt.SaveAs outFile, _
                         swSaveAsVersion_e.swSaveAsCurrentVersion, _
                         swSaveAsOptions_e.swSaveAsOptions_Silent Or swSaveAsOptions_e.swSaveAsOptions_Copy, _
                         exportData, errNum, warnNum
        Else
            ' Fallback: use SaveAs3 with the last used STL settings
            swModel.SaveAs3 outFile, 0, swSaveAsOptions_e.swSaveAsOptions_Silent
        End If
        
        If errNum = 0 Then
            savedCount = savedCount + 1
        End If
    Next i
    
    ' Display completion message
    MsgBox "Export complete! " & savedCount & " of " & n & _
           " configurations exported to:" & vbCrLf & outFolder, vbInformation

End Sub

' ------------------------------------------------------------------------
' CreateFolderIfNeeded
' Creates the specified folder if it does not already exist.
' ------------------------------------------------------------------------
Private Sub CreateFolderIfNeeded(ByVal folderPath As String)
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
End Sub

' ------------------------------------------------------------------------
' SanitizeFileName
' Replaces invalid Windows filename characters with underscores.
' ------------------------------------------------------------------------
Private Function SanitizeFileName(ByVal s As String) As String
    Dim badChars As Variant
    badChars = Array("<", ">", ":", """", "/", "\", "|", "?", "*")
    Dim i As Long
    For i = LBound(badChars) To UBound(badChars)
        s = Replace$(s, CStr(badChars(i)), "_")
    Next i
    s = Trim$(s)
    SanitizeFileName = s
End Function


