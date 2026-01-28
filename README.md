# SolidWorks DWG Layer Exporter & Image Generator

A SolidWorks VBA macro that batch processes `.dwg` files to:
1.  **Generate Preview Images:** Creates a high-quality PNG image of the drawing with *all layers visible*.
2.  **Filter Layers:** Creates a clean copy of the DWG with only specific layers visible (default: "0" and "ETCH").

## Features

- **Batch Processing:** Select a folder to process all `.dwg` files within it.
- **Visual Verification:** Automatically exports a PNG preview (in `DWG_Images` folder) before hiding layers, perfect for quick visual checks.
- **Layer Filtering:** Logic to keep "0" and "ETCH" layers visible while hiding others for the final production DWG (saved in `Filtered_DWGs`).
- **Non-Destructive:** Original files are left untouched.

## Installation & Usage

1.  **Open SolidWorks**.
2.  **Create the Macro:**
    *   Go to **Tools > Macro > New...**
    *   Save as `DWGLayerExporter.swp`.
    *   Copy the code from `DWGLayerExporter_FinalVersion1.bas` into the macro editor.
    *   Save and close the editor.
3.  **Run the Macro:**
    *   Go to **Tools > Macro > Run...** and select your macro.
    *   **Select Folder:** A dialog will ask for the folder containing your source DWG files.
4.  **Review Output:**
    *   The macro creates two new subfolders in your selected directory:
        *   `\DWG_Images\`: Contains PNG images of your parts (all layers visible).
        *   `\Filtered_DWGs\`: Contains the final DWG files (only "0" and "ETCH" layers visible).

## Configuration

### Changing Layer Rules
To change which layers are kept visible, edit this section in `DWGLayerExporter_FinalVersion1.bas`:

```vba
' Logic: If it's 0 or contains ETCH, Keep it visible. Otherwise, HIDE IT.
If curName = "0" Or InStr(curName, "ETCH") > 0 Then
    swLayer.Visible = True
Else
    swLayer.Visible = False
End If
```

## Files

| File | Description |
|------|-------------|
| `DWGLayerExporter_FinalVersion1.bas` | The complete VBA macro source code. |
| `README.md` | This documentation. |

## Requirements

- SolidWorks
- Windows OS

## License

MIT License
