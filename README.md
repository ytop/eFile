# Self-Extracted File Formatter

A simple file packing/unpacking system that packs all text and binary files from a directory (recursively) into a single text file, which can then be extracted back to the original directory structure.

## Components

1. **pack_files.py** - Python script to pack files
2. **extract_files.vba** - MS Access VBA script to extract files

## File Format Specification

The packed file uses a simple text-based format with the following structure:

```
================================================================================
SELF-EXTRACTED FILE ARCHIVE
Source Directory: <directory_name>
================================================================================

<<<FILE_START>>>
PATH: <relative/path/to/file.ext>
SIZE: <file_size_in_bytes>
ENCODING: base64
<<<CONTENT_START>>>
<base64_encoded_content>
<<<CONTENT_END>>>
<<<FILE_END>>>

<<<FILE_START>>>
PATH: <another/file.txt>
...
<<<FILE_END>>>

================================================================================
TOTAL FILES PACKED: <count>
================================================================================
```

### Format Details

- **Markers**: Uses clear text markers like `<<<FILE_START>>>`, `<<<CONTENT_START>>>`, etc.
- **Encoding**: All files are base64-encoded to safely store binary content in text format
- **Paths**: Relative paths using forward slashes (/) for cross-platform compatibility
- **Line Length**: Base64 content is wrapped at 76 characters per line (standard)

## Usage

### Packing Files (Python)

**Requirements**: Python 3.6+

**Basic Usage**:
```bash
python pack_files.py <source_directory> <output_file>
```

**Example**:
```bash
python pack_files.py ./myproject packed_files.txt
```

**Programmatic Usage**:
```python
from pack_files import pack_directory

# Pack all files from a directory
files_count = pack_directory("./myproject", "packed_files.txt")
print(f"Packed {files_count} files")
```

### Extracting Files (MS Access VBA)

**Requirements**: Microsoft Access 2010 or later

**Steps**:

1. Open Microsoft Access
2. Create a new blank database or open an existing one
3. Press `Alt+F11` to open the VBA Editor
4. Go to `Insert` → `Module` to create a new module
5. Copy the contents of `extract_files.vba` into the module
6. Modify the example paths in `Example_ExtractPackedFiles()` or call `ExtractFiles()` directly:

```vba
' From VBA Immediate Window (Ctrl+G) or a button click event
Call ExtractFiles("C:\temp\packed_files.txt", "C:\temp\extracted")
```

**Alternative - Using the Example Function**:

1. Edit the paths in `Example_ExtractPackedFiles()` subroutine
2. Run it from the VBA editor (F5) or call it from a form button

## Features

### Pack Function (Python)
- ✓ Recursively processes all subdirectories
- ✓ Handles both text and binary files
- ✓ Base64 encoding for safe text storage
- ✓ Preserves relative directory structure
- ✓ Progress reporting during packing
- ✓ Error handling for inaccessible files
- ✓ Cross-platform path handling

### Extract Function (VBA)
- ✓ Recreates original directory structure
- ✓ Base64 decoding using MSXML2
- ✓ Handles binary files correctly
- ✓ Creates missing directories automatically
- ✓ Windows path conversion (/ to \)
- ✓ Progress reporting via Debug.Print
- ✓ User-friendly message boxes

## Example Workflow

### Pack files:
```bash
# Pack your project files
python pack_files.py ./my_project project_backup.txt
```

### Extract files:
```vba
' In MS Access VBA
Call ExtractFiles("C:\backups\project_backup.txt", "C:\restored\my_project")
```

## Limitations

1. **File Size**: Very large files will create very large packed files (base64 encoding increases size by ~33%)
2. **Text Format**: The packed file must remain as text; do not edit manually
3. **Platform**: Extract function is Windows-only (MS Access VBA)
4. **Permissions**: Both pack and extract require appropriate file system permissions

## File Size Considerations

Base64 encoding increases file size by approximately 33%. For example:
- 1 MB of files → ~1.33 MB packed file
- 100 MB of files → ~133 MB packed file

For very large directories, consider:
- Filtering specific file types
- Splitting into multiple packed files
- Using compression before packing

## Error Handling

### Pack Function
- Validates source directory exists
- Skips files that cannot be read
- Reports warnings for skipped files
- Continues processing on individual file errors

### Extract Function
- Validates packed file exists
- Creates output directory if missing
- Reports extraction progress
- Shows summary message box when complete

## Technical Notes

### Python Dependencies
- Uses only Python standard library (no external packages required)
- `pathlib` for cross-platform path handling
- `base64` for encoding
- `os` for directory traversal

### VBA Dependencies
- `Scripting.FileSystemObject` for file operations
- `MSXML2.DOMDocument` for base64 decoding
- `ADODB.Stream` for binary file writing

## License

This is a simple utility script provided as-is for file management purposes.
