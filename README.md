# SolidWorks DWG Layer Exporter

A SolidWorks VBA macro that batch exports drawing files (.SLDDRW) to DWG format with selective layer visibility.

## Features

- **Batch Processing**: Export multiple drawings at once from a folder
- **Layer Filtering**: Keep only specified layers (default: "Drawing View 1" and "ETCH")
- **Multi-Sheet Support**: Handles drawings with multiple sheets
  - Sheet 1 → `filename.dwg`
  - Sheet 2 → `filenameFLO.dwg`
- **Bend Line Hiding**: Automatically hides sheet metal bend lines in exports
- **Non-Destructive**: Original drawings are never modified

## Quick Start

### Installation

1. Open SolidWorks
2. Go to **Tools → Macro → New** and save as `DWGLayerExporter.swp`
3. In the VBA Editor, go to **Tools → References** and enable:
   - SolidWorks Type Library
   - SolidWorks Constant Type Library
4. Copy all code from `DWGLayerExporter_Simple.bas` into the module
5. Save and close VBA Editor

### Usage

1. Run the macro: **Tools → Macro → Run → DWGLayerExporter.swp**
2. Select the folder containing your `.SLDDRW` files
3. Select the output folder for DWG files
4. Confirm the export settings
5. Wait for processing to complete

## Configuration

### Changing Default Layers

Edit these lines at the top of the macro:

```vba
Private Const KEEP_LAYER_1 As String = "Drawing View 1"
Private Const KEEP_LAYER_2 As String = "ETCH"
```

To add more layers, add additional constants and update the `LayerShouldBeKept` function:

```vba
Private Function LayerShouldBeKept(layerName As String) As Boolean
    LayerShouldBeKept = (layerName = KEEP_LAYER_1) Or _
                        (layerName = KEEP_LAYER_2) Or _
                        (layerName = "Your New Layer")
End Function
```

## Files

| File | Description |
|------|-------------|
| `DWGLayerExporter_Simple.bas` | Main macro code (recommended) |
| `DWGLayerExporter.swp.bas` | Full version with layer selection GUI |
| `LayerSelectionForm.frm` | UserForm code for GUI version |
| `INSTALLATION_INSTRUCTIONS.txt` | Detailed setup guide |

## Requirements

- SolidWorks (tested on 2020+)
- Windows OS

## Troubleshooting

**"SolidWorks is not running"**
- Make sure SolidWorks is open before running the macro

**"No SLDDRW files found"**
- Verify the folder contains `.SLDDRW` files

**Some files fail to export**
- Files may be corrupted or currently open
- Check the VBA Immediate Window (View → Immediate Window) for details

## License

MIT License - Feel free to use and modify.
