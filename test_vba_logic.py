"""
Test script to verify the VBA extraction logic
This simulates what the VBA script does using Python
"""
import base64
import os
from pathlib import Path


def extract_files_test(packed_file, output_dir):
    """
    Python version of the VBA extract logic for testing
    """
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    with open(packed_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    current_file_path = None
    current_encoding = None
    content_lines = []
    in_content = False
    files_extracted = 0

    for line in lines:
        line = line.strip()

        if "<<<FILE_START>>>" in line:
            current_file_path = None
            current_encoding = None
            content_lines = []
            in_content = False

        elif line.startswith("PATH:"):
            current_file_path = line.split(":", 1)[1].strip()

        elif line.startswith("ENCODING:"):
            current_encoding = line.split(":", 1)[1].strip()

        elif "<<<CONTENT_START>>>" in line:
            in_content = True
            content_lines = []

        elif "<<<CONTENT_END>>>" in line:
            in_content = False

        elif "<<<FILE_END>>>" in line:
            if current_file_path:
                # Decode and write file (allow empty files)
                encoded_content = "".join(content_lines)
                if encoded_content:
                    decoded_content = base64.b64decode(encoded_content)
                else:
                    decoded_content = b""  # Empty file

                full_path = output_path / current_file_path
                full_path.parent.mkdir(parents=True, exist_ok=True)

                with open(full_path, 'wb') as out_file:
                    out_file.write(decoded_content)

                print(f"Extracted: {current_file_path}")
                files_extracted += 1

        elif in_content:
            content_lines.append(line)

    print(f"\nSuccessfully extracted {files_extracted} files to {output_dir}")
    return files_extracted


if __name__ == "__main__":
    # Test extraction
    extract_files_test("test_output_new.txt", "test_extracted")

    # Verify extracted files match originals
    print("\n=== Verification ===")
    test_files = [
        ("test_data/empty_file.txt", "test_extracted/empty_file.txt"),
        ("test_data/file1.txt", "test_extracted/file1.txt"),
        ("test_data/subdir1/file2.txt", "test_extracted/subdir1/file2.txt"),
        ("test_data/subdir1/subdir2/file3.txt", "test_extracted/subdir1/subdir2/file3.txt")
    ]

    all_match = True
    for original, extracted in test_files:
        with open(original, 'rb') as f1, open(extracted, 'rb') as f2:
            original_content = f1.read()
            extracted_content = f2.read()

            if original_content == extracted_content:
                print(f"✓ {original} matches")
            else:
                print(f"✗ {original} does NOT match")
                all_match = False

    if all_match:
        print("\n✓✓✓ All files extracted correctly! VBA logic verified.")
    else:
        print("\n✗✗✗ Some files do not match!")
