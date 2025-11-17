# Quick Start Guide

## Packing Files (Python)

```bash
# Pack all files in a directory
python3 pack_files.py <source_directory> <output_file>

# Example:
python3 pack_files.py ./my_project backup.txt
```

## Extracting Files (MS Access VBA)

### Method 1: Direct Function Call

1. Open Microsoft Access
2. Press `Alt+F11` (VBA Editor)
3. Insert → Module
4. Copy contents of `extract_files.vba`
5. Press `Ctrl+G` (Immediate Window)
6. Type and run:

```vba
Call ExtractFiles("C:\path\to\backup.txt", "C:\path\to\output_folder")
```

### Method 2: Use Example Function

1. Follow steps 1-4 above
2. Edit the `Example_ExtractPackedFiles()` function:

```vba
Public Sub Example_ExtractPackedFiles()
    Dim packedFile As String
    Dim outputDirectory As String

    ' Change these paths
    packedFile = "C:\backup.txt"
    outputDirectory = "C:\restored"

    Call ExtractFiles(packedFile, outputDirectory)
End Sub
```

3. Press `F5` to run

## File Format Summary

The packed file uses markers to separate files:

```
<<<FILE_START>>>
PATH: path/to/file.txt
SIZE: 123
ENCODING: base64
<<<CONTENT_START>>>
<base64 content>
<<<CONTENT_END>>>
<<<FILE_END>>>
```

## Testing

Run the included test:

```bash
python3 test_vba_logic.py
```

This verifies the pack/extract logic works correctly.
