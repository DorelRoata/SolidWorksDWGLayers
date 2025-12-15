' ============================================================================
' LayerSelectionForm - UserForm for selecting which layers to export
' ============================================================================
' This form displays all available layers and lets user select which to keep
' ============================================================================

' === FORM DESIGN INSTRUCTIONS ===
' Create a new UserForm named "LayerSelectionForm" with these controls:
'
' 1. ListBox:
'    - Name: lstLayers
'    - MultiSelect: fmMultiSelectMulti
'    - Top: 40, Left: 10, Width: 260, Height: 200
'
' 2. Label:
'    - Name: lblInstructions
'    - Caption: "Select layers to KEEP in the exported DWG files:"
'    - Top: 10, Left: 10, Width: 260
'
' 3. CommandButton (OK):
'    - Name: btnOK
'    - Caption: "Export DWG Files"
'    - Top: 250, Left: 100, Width: 80, Height: 25
'
' 4. CommandButton (Cancel):
'    - Name: btnCancel
'    - Caption: "Cancel"
'    - Top: 250, Left: 190, Width: 80, Height: 25
'
' 5. CommandButton (Select Defaults):
'    - Name: btnDefaults
'    - Caption: "Select Defaults"
'    - Top: 250, Left: 10, Width: 80, Height: 25
'
' 6. Label (Status):
'    - Name: lblStatus
'    - Caption: ""
'    - Top: 280, Left: 10, Width: 260
'
' Form Properties:
'    - Caption: "DWG Layer Exporter"
'    - Width: 300, Height: 320

Option Explicit

' Default layers to select
Private Const DEFAULT_LAYER_1 As String = "Drawing View 1"
Private Const DEFAULT_LAYER_2 As String = "ETCH"

Private Sub UserForm_Initialize()
    ' Populate the layer list from the first drawing file
    Dim allLayers As Collection
    Dim filePath As Variant
    Dim layerName As Variant
    Dim foundLayers As New Collection

    lblStatus.Caption = "Loading layers..."
    Me.Repaint

    ' Scan first few files to get all unique layers
    Dim fileCount As Long
    fileCount = 0

    For Each filePath In FilesToProcess
        If fileCount >= 3 Then Exit For ' Only scan first 3 files for speed

        Dim layers As Collection
        Set layers = GetDrawingLayers(CStr(filePath))

        For Each layerName In layers
            On Error Resume Next
            foundLayers.Add CStr(layerName), CStr(layerName)
            On Error GoTo 0
        Next layerName

        fileCount = fileCount + 1
    Next filePath

    ' Add layers to listbox
    lstLayers.Clear
    For Each layerName In foundLayers
        lstLayers.AddItem CStr(layerName)
    Next layerName

    ' Pre-select default layers
    SelectDefaultLayers

    lblStatus.Caption = "Found " & lstLayers.ListCount & " layers. Select layers to keep."
End Sub

Private Sub btnDefaults_Click()
    SelectDefaultLayers
End Sub

Private Sub SelectDefaultLayers()
    Dim i As Long

    ' First deselect all
    For i = 0 To lstLayers.ListCount - 1
        lstLayers.Selected(i) = False
    Next i

    ' Select default layers
    For i = 0 To lstLayers.ListCount - 1
        If lstLayers.List(i) = DEFAULT_LAYER_1 Or _
           lstLayers.List(i) = DEFAULT_LAYER_2 Then
            lstLayers.Selected(i) = True
        End If
    Next i
End Sub

Private Sub btnOK_Click()
    ' Collect selected layers
    Set SelectedLayers = New Collection

    Dim i As Long
    For i = 0 To lstLayers.ListCount - 1
        If lstLayers.Selected(i) Then
            SelectedLayers.Add lstLayers.List(i)
        End If
    Next i

    If SelectedLayers.Count = 0 Then
        MsgBox "Please select at least one layer to keep!", vbExclamation, "No Layers Selected"
        Exit Sub
    End If

    ' Confirm and process
    Dim msg As String
    msg = "Ready to export " & FilesToProcess.Count & " drawing(s)" & vbCrLf
    msg = msg & "Keeping " & SelectedLayers.Count & " layer(s):" & vbCrLf & vbCrLf

    Dim layerName As Variant
    For Each layerName In SelectedLayers
        msg = msg & "  - " & layerName & vbCrLf
    Next layerName

    msg = msg & vbCrLf & "Output folder: " & OutputFolder

    If MsgBox(msg, vbYesNo + vbQuestion, "Confirm Export") = vbYes Then
        Me.Hide
        ProcessDrawings
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
