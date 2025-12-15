' ============================================================================
' DWG Layer Exporter for SolidWorks
' ============================================================================
' This macro exports SolidWorks drawings to DWG with only selected layers visible
' Default layers: "Drawing View 1" and "ETCH"
' ============================================================================

Option Explicit

' Global variables
Public swApp As SldWorks.SldWorks
Public SelectedLayers As Collection
Public OutputFolder As String
Public FilesToProcess As Collection

' Main entry point
Sub main()
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        MsgBox "SolidWorks is not running!", vbCritical, "Error"
        Exit Sub
    End If

    ' Initialize collections
    Set FilesToProcess = New Collection
    Set SelectedLayers = New Collection

    ' Step 1: Select drawing files
    If Not SelectDrawingFiles() Then
        Exit Sub
    End If

    ' Step 2: Select output folder
    If Not SelectOutputFolder() Then
        Exit Sub
    End If

    ' Step 3: Show layer selection dialog
    LayerSelectionForm.Show

    ' Processing continues in the form's OK button click event
End Sub

' Function to select multiple drawing files
Private Function SelectDrawingFiles() As Boolean
    Dim fd As Object
    Dim selectedItems As Variant
    Dim i As Long

    Set fd = CreateObject("Shell.Application").BrowseForFolder(0, "Select folder containing SLDDRW files", 0)

    If fd Is Nothing Then
        SelectDrawingFiles = False
        Exit Function
    End If

    Dim folderPath As String
    folderPath = fd.Self.Path

    ' Use FileSystemObject to get all SLDDRW files
    Dim fso As Object
    Dim folder As Object
    Dim file As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    Dim fileCount As Long
    fileCount = 0

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "slddrw" Then
            FilesToProcess.Add file.Path
            fileCount = fileCount + 1
        End If
    Next file

    If fileCount = 0 Then
        MsgBox "No SLDDRW files found in the selected folder!", vbExclamation, "No Files"
        SelectDrawingFiles = False
        Exit Function
    End If

    Dim response As VbMsgBoxResult
    response = MsgBox("Found " & fileCount & " drawing file(s). Continue?", vbYesNo + vbQuestion, "Files Found")

    SelectDrawingFiles = (response = vbYes)
End Function

' Function to select output folder
Private Function SelectOutputFolder() As Boolean
    Dim fd As Object

    Set fd = CreateObject("Shell.Application").BrowseForFolder(0, "Select OUTPUT folder for DWG files", 0)

    If fd Is Nothing Then
        SelectOutputFolder = False
        Exit Function
    End If

    OutputFolder = fd.Self.Path

    ' Ensure trailing backslash
    If Right(OutputFolder, 1) <> "\" Then
        OutputFolder = OutputFolder & "\"
    End If

    SelectOutputFolder = True
End Function

' Function to get all layers from a drawing
Public Function GetDrawingLayers(drawingPath As String) As Collection
    Dim layers As New Collection
    Dim swModel As SldWorks.ModelDoc2
    Dim swDraw As SldWorks.DrawingDoc
    Dim swLayerMgr As SldWorks.LayerMgr
    Dim layerNames As Variant
    Dim i As Long
    Dim errors As Long
    Dim warnings As Long

    ' Open the drawing silently
    Set swModel = swApp.OpenDoc6(drawingPath, swDocDRAWING, swOpenDocOptions_Silent, "", errors, warnings)

    If swModel Is Nothing Then
        Set GetDrawingLayers = layers
        Exit Function
    End If

    Set swDraw = swModel
    Set swLayerMgr = swModel.GetLayerManager

    layerNames = swLayerMgr.GetLayerList

    If Not IsEmpty(layerNames) Then
        For i = LBound(layerNames) To UBound(layerNames)
            On Error Resume Next
            layers.Add CStr(layerNames(i)), CStr(layerNames(i))
            On Error GoTo 0
        Next i
    End If

    swApp.CloseDoc swModel.GetPathName

    Set GetDrawingLayers = layers
End Function

' Main processing function - called from the form
Public Sub ProcessDrawings()
    Dim i As Long
    Dim filePath As String
    Dim fileName As String
    Dim swModel As SldWorks.ModelDoc2
    Dim swDraw As SldWorks.DrawingDoc
    Dim swLayerMgr As SldWorks.LayerMgr
    Dim swLayer As SldWorks.Layer
    Dim layerNames As Variant
    Dim j As Long
    Dim errors As Long
    Dim warnings As Long
    Dim exportOptions As Long
    Dim fso As Object
    Dim successCount As Long
    Dim failCount As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    successCount = 0
    failCount = 0

    ' Process each file
    For i = 1 To FilesToProcess.Count
        filePath = FilesToProcess(i)
        fileName = fso.GetBaseName(filePath)

        ' Open the drawing
        Set swModel = swApp.OpenDoc6(filePath, swDocDRAWING, swOpenDocOptions_Silent, "", errors, warnings)

        If swModel Is Nothing Then
            failCount = failCount + 1
            Debug.Print "Failed to open: " & filePath
        Else
            Set swDraw = swModel
            Set swLayerMgr = swModel.GetLayerManager

            ' Get all layers
            layerNames = swLayerMgr.GetLayerList

            If Not IsEmpty(layerNames) Then
                ' Hide all layers first
                For j = LBound(layerNames) To UBound(layerNames)
                    Set swLayer = swLayerMgr.GetLayer(CStr(layerNames(j)))
                    If Not swLayer Is Nothing Then
                        swLayer.Visible = False
                    End If
                Next j

                ' Show only selected layers
                Dim layerName As Variant
                For Each layerName In SelectedLayers
                    Set swLayer = swLayerMgr.GetLayer(CStr(layerName))
                    If Not swLayer Is Nothing Then
                        swLayer.Visible = True
                    End If
                Next layerName
            End If

            ' Rebuild to apply layer changes
            swModel.ForceRebuild3 False

            ' Export to DWG
            Dim outputPath As String
            outputPath = OutputFolder & fileName & ".dwg"

            ' Use SaveAs with DWG format
            Dim saveErrors As Long
            Dim saveWarnings As Long
            Dim saveResult As Boolean

            saveResult = swModel.Extension.SaveAs(outputPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, saveErrors, saveWarnings)

            If saveResult Then
                successCount = successCount + 1
                Debug.Print "Exported: " & outputPath
            Else
                failCount = failCount + 1
                Debug.Print "Failed to export: " & outputPath & " (Error: " & saveErrors & ")"
            End If

            ' Close without saving changes to original
            swApp.CloseDoc swModel.GetPathName
        End If
    Next i

    ' Show summary
    MsgBox "Export Complete!" & vbCrLf & vbCrLf & _
           "Successfully exported: " & successCount & " file(s)" & vbCrLf & _
           "Failed: " & failCount & " file(s)" & vbCrLf & vbCrLf & _
           "Output folder: " & OutputFolder, _
           vbInformation, "DWG Export Complete"
End Sub
