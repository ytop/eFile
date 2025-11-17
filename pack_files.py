"""
Self-Extracted File Formatter - Pack Function
Packs all files from a directory (recursively) into a single text file.
"""

import os
import base64
from pathlib import Path


def pack_directory(source_dir, output_file):
    """
    Pack all files from source_dir into a single text file.

    Args:
        source_dir: Directory to pack (will process recursively)
        output_file: Output text file path
    """
    source_path = Path(source_dir).resolve()
    output_path = Path(output_file).resolve()

    if not source_path.exists():
        raise ValueError(f"Source directory does not exist: {source_dir}")

    if not source_path.is_dir():
        raise ValueError(f"Source path is not a directory: {source_dir}")

    # Check if output file would be inside source directory
    try:
        output_path.relative_to(source_path)
        print(f"Warning: Output file is inside source directory and will be excluded from packing")
    except ValueError:
        pass  # Output file is outside source directory, which is fine

    with open(output_file, 'w', encoding='utf-8') as out:
        # Write header
        out.write("=" * 80 + "\n")
        out.write("SELF-EXTRACTED FILE ARCHIVE\n")
        out.write(f"Source Directory: {source_path.name}\n")
        out.write("=" * 80 + "\n\n")

        # Collect all files
        files_packed = 0
        files_skipped = 0

        for root, dirs, files in os.walk(source_path, followlinks=False):
            for filename in files:
                file_path = Path(root) / filename

                # Skip symlinks
                if file_path.is_symlink():
                    print(f"Skipped (symlink): {file_path.relative_to(source_path)}")
                    files_skipped += 1
                    continue

                # Skip the output file itself
                if file_path.resolve() == output_path:
                    print(f"Skipped (output file): {file_path.relative_to(source_path)}")
                    files_skipped += 1
                    continue

                # Get relative path from source directory
                try:
                    relative_path = file_path.relative_to(source_path)
                except ValueError:
                    continue

                # Read file content
                try:
                    with open(file_path, 'rb') as f:
                        content = f.read()

                    # Encode to base64 for safe text storage
                    encoded_content = base64.b64encode(content).decode('ascii')

                    # Write file entry
                    out.write("<<<FILE_START>>>\n")
                    out.write(f"PATH: {relative_path.as_posix()}\n")
                    out.write(f"SIZE: {len(content)}\n")
                    out.write(f"ENCODING: base64\n")
                    out.write("<<<CONTENT_START>>>\n")

                    # Write content in chunks of 76 characters (standard base64 line length)
                    # Handle empty files (no content to write)
                    if encoded_content:
                        for i in range(0, len(encoded_content), 76):
                            out.write(encoded_content[i:i+76] + "\n")

                    out.write("<<<CONTENT_END>>>\n")
                    out.write("<<<FILE_END>>>\n\n")

                    files_packed += 1
                    print(f"Packed: {relative_path}")

                except PermissionError:
                    print(f"Warning: Permission denied - {relative_path}")
                    files_skipped += 1
                except FileNotFoundError:
                    print(f"Warning: File not found (may have been deleted) - {relative_path}")
                    files_skipped += 1
                except Exception as e:
                    print(f"Warning: Could not pack {relative_path}: {type(e).__name__}: {e}")
                    files_skipped += 1

        # Write footer
        out.write("=" * 80 + "\n")
        out.write(f"TOTAL FILES PACKED: {files_packed}\n")
        out.write("=" * 80 + "\n")

    print(f"\n✓ Successfully packed {files_packed} files into {output_file}")
    if files_skipped > 0:
        print(f"⚠ Skipped {files_skipped} files")
    return files_packed


def main():
    """Example usage"""
    import sys

    if len(sys.argv) < 3:
        print("Usage: python pack_files.py <source_directory> <output_file>")
        print("Example: python pack_files.py ./myproject packed_files.txt")
        sys.exit(1)

    source_dir = sys.argv[1]
    output_file = sys.argv[2]

    try:
        pack_directory(source_dir, output_file)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
