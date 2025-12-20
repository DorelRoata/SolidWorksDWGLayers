' ============================================================================
' DWG Layer Exporter - SIMPLE VERSION (No UserForm Required)
' ============================================================================
' Exports SolidWorks drawings to DWG with only selected layers visible
' - Sheet 1 exports as: filename.dwg
' - Sheet 2 exports as: filenameFLO.dwg
' - Bendlines optionally hidden (set HIDE_BENDLINES = True/False)
' ============================================================================

Option Explicit

' ============================================================================
' CONFIGURE YOUR LAYERS HERE - Edit these to change which layers are kept
' ============================================================================
Private Const KEEP_LAYER_1 As String = "Drawing View 1"
Private Const KEEP_LAYER_2 As String = "ETCH"
' Add more as needed (also update the LayerShouldBeKept function below)

' Bendline hiding - Set to True to hide bendlines, False to skip
Private Const HIDE_BENDLINES As Boolean = True
' ============================================================================

Private swApp As SldWorks.SldWorks

Sub main()
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        MsgBox "SolidWorks is not running!", vbCritical, "Error"
        Exit Sub
    End If

    ' Select source folder
    Dim sourceFolder As String
    sourceFolder = BrowseForFolder("Select folder containing SLDDRW files")
    If sourceFolder = "" Then Exit Sub

    ' Select output folder
    Dim outputFolder As String
    outputFolder = BrowseForFolder("Select OUTPUT folder for DWG files")
    If outputFolder = "" Then Exit Sub

    ' Ensure trailing backslash
    If Right(outputFolder, 1) <> "\" Then outputFolder = outputFolder & "\"

    ' Get all SLDDRW files
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folder As Object
    Set folder = fso.GetFolder(sourceFolder)

    Dim files As New Collection
    Dim file As Object

    For Each file In folder.files
        If LCase(fso.GetExtensionName(file.Name)) = "slddrw" Then
            files.Add file.Path
        End If
    Next file

    If files.Count = 0 Then
        MsgBox "No SLDDRW files found in:" & vbCrLf & sourceFolder, vbExclamation, "No Files"
        Exit Sub
    End If

    ' Confirm
    Dim msg As String
    msg = "Found " & files.Count & " drawing file(s)" & vbCrLf & vbCrLf
    msg = msg & "Layers to KEEP:" & vbCrLf
    msg = msg & "  - " & KEEP_LAYER_1 & vbCrLf
    msg = msg & "  - " & KEEP_LAYER_2 & vbCrLf & vbCrLf
    msg = msg & "Sheet naming:" & vbCrLf
    msg = msg & "  - Sheet 1 -> filename.dwg" & vbCrLf
    msg = msg & "  - Sheet 2 -> filenameFLO.dwg" & vbCrLf & vbCrLf
    If HIDE_BENDLINES Then
        msg = msg & "Bendlines will be hidden" & vbCrLf & vbCrLf
    End If
    msg = msg & "Output folder:" & vbCrLf & outputFolder & vbCrLf & vbCrLf
    msg = msg & "Continue with export?"

    If MsgBox(msg, vbYesNo + vbQuestion, "Confirm Export") <> vbYes Then
        Exit Sub
    End If

    ' Process files
    Dim successCount As Long
    Dim failCount As Long
    Dim filePath As Variant

    For Each filePath In files
        Dim result As Long
        result = ProcessDrawing(CStr(filePath), outputFolder, fso)
        successCount = successCount + result
        If result = 0 Then failCount = failCount + 1
    Next filePath

    ' Show summary
    MsgBox "Export Complete!" & vbCrLf & vbCrLf & _
           "Successfully exported: " & successCount & " DWG file(s)" & vbCrLf & _
           "Drawings with errors: " & failCount & vbCrLf & vbCrLf & _
           "Output folder: " & outputFolder, _
           vbInformation, "DWG Export Complete"
End Sub

' Check if a layer should be kept visible
Private Function LayerShouldBeKept(layerName As String) As Boolean
    ' Add more conditions here if you have more layers to keep
    LayerShouldBeKept = (layerName = KEEP_LAYER_1) Or _
                        (layerName = KEEP_LAYER_2)
End Function

' Process a single drawing file - returns number of successful exports
Private Function ProcessDrawing(filePath As String, outputFolder As String, fso As Object) As Long
    Dim swModel As SldWorks.ModelDoc2
    Dim swDraw As SldWorks.DrawingDoc
    Dim swLayerMgr As SldWorks.LayerMgr
    Dim swLayer As SldWorks.Layer
    Dim layerNames As Variant
    Dim i As Long
    Dim errors As Long
    Dim warnings As Long
    Dim sheetNames As Variant
    Dim sheetCount As Long
    Dim exportCount As Long

    ProcessDrawing = 0
    exportCount = 0

    ' Open the drawing
    Set swModel = swApp.OpenDoc6(filePath, swDocDRAWING, swOpenDocOptions_Silent, "", errors, warnings)

    If swModel Is Nothing Then
        Debug.Print "Failed to open: " & filePath
        Exit Function
    End If

    Set swDraw = swModel
    Set swLayerMgr = swModel.GetLayerManager

    ' Get sheet names
    sheetNames = swDraw.GetSheetNames
    sheetCount = UBound(sheetNames) - LBound(sheetNames) + 1

    Dim fileName As String
    fileName = fso.GetBaseName(filePath)

    ' Process each sheet
    Dim sheetIndex As Long
    For sheetIndex = LBound(sheetNames) To UBound(sheetNames)
        Dim sheetName As String
        sheetName = sheetNames(sheetIndex)

        ' Activate this sheet
        swDraw.ActivateSheet sheetName

        ' Set layer visibility
        layerNames = swLayerMgr.GetLayerList
        If Not IsEmpty(layerNames) Then
            For i = LBound(layerNames) To UBound(layerNames)
                Set swLayer = swLayerMgr.GetLayer(CStr(layerNames(i)))
                If Not swLayer Is Nothing Then
                    swLayer.Visible = LayerShouldBeKept(CStr(layerNames(i)))
                End If
            Next i
        End If

        ' Hide bendlines if enabled
        If HIDE_BENDLINES Then
            HideBendLines swDraw
        End If

        ' Rebuild to apply changes
        swModel.ForceRebuild3 False

        ' Determine output filename based on sheet number
        Dim outputPath As String
        Dim sheetNum As Long
        sheetNum = sheetIndex - LBound(sheetNames) + 1 ' 1-based sheet number

        If sheetNum = 1 Then
            ' First sheet: filename.dwg
            outputPath = outputFolder & fileName & ".dwg"
        ElseIf sheetNum = 2 Then
            ' Second sheet: filenameFLO.dwg
            outputPath = outputFolder & fileName & "FLO.dwg"
        Else
            ' Additional sheets: filename_Sheet3.dwg, etc.
            outputPath = outputFolder & fileName & "_Sheet" & sheetNum & ".dwg"
        End If

        ' Export current sheet to DWG
        Dim saveErrors As Long
        Dim saveWarnings As Long
        Dim saveResult As Boolean

        ' Use sheet-specific export
        saveResult = ExportSheetToDWG(swModel, swDraw, sheetName, outputPath)

        If saveResult Then
            exportCount = exportCount + 1
            Debug.Print "Exported: " & outputPath
        Else
            Debug.Print "Failed to export sheet " & sheetNum & " of: " & filePath
        End If
    Next sheetIndex

    ' Close without saving changes to original
    swApp.CloseDoc swModel.GetPathName

    ProcessDrawing = exportCount
End Function

' Export a specific sheet to DWG
Private Function ExportSheetToDWG(swModel As SldWorks.ModelDoc2, swDraw As SldWorks.DrawingDoc, sheetName As String, outputPath As String) As Boolean
    Dim swExportData As SldWorks.ExportPdfData
    Dim saveErrors As Long
    Dim saveWarnings As Long

    ' Make sure the sheet is active
    swDraw.ActivateSheet sheetName

    ' For DWG export, we need to use SaveAs with proper options
    ' swSaveAsOptions_Silent = 1
    ' The active sheet will be exported

    ExportSheetToDWG = swModel.Extension.SaveAs(outputPath, swSaveAsCurrentVersion, _
                                                 swSaveAsOptions_Silent, Nothing, _
                                                 saveErrors, saveWarnings)
End Function

' Hide bendlines in drawing views
Private Sub HideBendLines(swDraw As SldWorks.DrawingDoc)
    On Error Resume Next

    Dim swView As SldWorks.View
    Dim swSheet As SldWorks.sheet
    Dim vSheets As Variant
    Dim vViews As Variant
    Dim i As Long, j As Long

    ' Get current sheet
    Set swSheet = swDraw.GetCurrentSheet

    ' Get all views on the sheet
    vViews = swSheet.GetViews

    If IsEmpty(vViews) Then Exit Sub

    For i = LBound(vViews) To UBound(vViews)
        Set swView = vViews(i)
        If Not swView Is Nothing Then
            ' Try to hide bend lines through the view
            ' Bend lines are typically sketch entities or annotations
            HideBendLinesInView swDraw, swView
        End If
    Next i

    On Error GoTo 0
End Sub

' Hide bend lines in a specific view
Private Sub HideBendLinesInView(swDraw As SldWorks.DrawingDoc, swView As SldWorks.View)
    On Error Resume Next

    ' Get the model from the view
    Dim swRefModel As SldWorks.ModelDoc2
    Set swRefModel = swView.ReferencedDocument

    If swRefModel Is Nothing Then Exit Sub

    ' For sheet metal parts, try to hide bend lines feature
    Dim swFeat As SldWorks.Feature
    Set swFeat = swRefModel.FirstFeature

    Do While Not swFeat Is Nothing
        Dim featName As String
        featName = LCase(swFeat.Name)

        ' Look for bend-related features
        If InStr(featName, "bend") > 0 Or _
           InStr(featName, "flat") > 0 Then
            ' Try to set visibility in drawing view
            swView.SetBendLineVisibility False
        End If

        Set swFeat = swFeat.GetNextFeature
    Loop

    ' Also try the direct method
    swView.SetBendLineVisibility False

    On Error GoTo 0
End Sub

' Browse for folder dialog
Private Function BrowseForFolder(prompt As String) As String
    Dim shell As Object
    Dim folder As Object

    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, prompt, 0)

    If folder Is Nothing Then
        BrowseForFolder = ""
    Else
        BrowseForFolder = folder.Self.Path
    End If
End Function
