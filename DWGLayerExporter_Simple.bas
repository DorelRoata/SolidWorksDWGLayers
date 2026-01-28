Option Explicit

Sub main()
    Dim swApp As Object
    Dim swModel As Object
    Dim swImportData As Object
    Dim swLayerMgr As Object
    Dim swLayer As Object
    Dim vLayers As Variant
    Dim i As Long
    Dim fileFolder As String, fileName As String, fullPath As String
    Dim outputFolder As String, imagesFolder As String
    Dim savePath As String, imagePath As String
    Dim errors As Long, warnings As Long
    
    Set swApp = Application.SldWorks
    
    ' 1. Select Folder
    fileFolder = GetFolder()
    If fileFolder = "" Then Exit Sub
    If Right(fileFolder, 1) <> "\" Then fileFolder = fileFolder & "\"
    
    ' Setup Output Folders
    outputFolder = fileFolder & "Filtered_DWGs\"
    If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder
    
    imagesFolder = fileFolder & "DWG_Images\"
    If Dir(imagesFolder, vbDirectory) = "" Then MkDir imagesFolder
    
    ' --- SET SOLIDWORKS SYSTEM EXPORT SETTING ---
    ' This ensures hidden layers are NOT included in the final DWG
    swApp.SetUserPreferenceToggle 108, False ' 108 = swDxfExportExportAllSheetsToOneFile
    
    fileName = Dir(fileFolder & "*.dwg")
    
    Do While fileName <> ""
        fullPath = fileFolder & fileName
        Set swImportData = swApp.GetImportFileData(fullPath)
        
        If Not swImportData Is Nothing Then
            Set swModel = swApp.LoadFile4(fullPath, "", swImportData, errors)
            
            If Not swModel Is Nothing Then
                Set swLayerMgr = swModel.GetLayerManager
                vLayers = swLayerMgr.GetLayerList
                
                If Not IsEmpty(vLayers) Then
                    
                    ' --- STEP 1: EXPORT PREVIEW IMAGE (ALL LAYERS VISIBLE) ---
                    ' Ensure all layers are visible for the screenshot
                    For i = 0 To UBound(vLayers)
                        Set swLayer = swLayerMgr.GetLayer(vLayers(i))
                        swLayer.Visible = True
                    Next i
                    
                    ' Zoom to fit and Save Image
                    swModel.ViewZoomtofit2
                    ' Replace .dwg extension with .png for the image filename
                    imagePath = imagesFolder & Left(fileName, InStrRev(fileName, ".") - 1) & ".png"
                    swModel.Extension.SaveAs imagePath, 0, 1, Nothing, errors, warnings
                    
                    ' --- STEP 2: FILTER LAYERS FOR DWG EXPORT ---
                    For i = 0 To UBound(vLayers)
                        Set swLayer = swLayerMgr.GetLayer(vLayers(i))
                        Dim curName As String
                        curName = UCase(Trim(vLayers(i)))
                        
                        ' Logic: If it's 0 or contains ETCH, Keep it visible. Otherwise, HIDE IT.
                        If curName = "0" Or InStr(curName, "ETCH") > 0 Then
                            swLayer.Visible = True
                        Else
                            swLayer.Visible = False
                        End If
                    Next i
                End If
                
                ' --- STEP 3: SAVE FILTERED DWG ---
                swModel.ViewZoomtofit2
                savePath = outputFolder & fileName
                swModel.Extension.SaveAs savePath, 0, 1, Nothing, errors, warnings
                
                swApp.CloseDoc swModel.GetTitle
            End If
        End If
        fileName = Dir
    Loop
    
    MsgBox "Batch Complete!" & vbCrLf & _
           "1. Images saved in 'DWG_Images'" & vbCrLf & _
           "2. Cleaned DWGs saved in 'Filtered_DWGs'"
End Sub

Function GetFolder() As String
    Dim fldr As Object
    Set fldr = CreateObject("Shell.Application").BrowseForFolder(0, "Select Folder", &H41)
    If Not fldr Is Nothing Then GetFolder = fldr.Items.Item.Path
End Function
