"""
Test edge cases for the pack/extract system
"""
import os
import sys
from pack_files import pack_directory
from test_vba_logic import extract_files_test

def test_edge_cases():
    print("=" * 80)
    print("EDGE CASE TESTING")
    print("=" * 80)

    # Create test directory
    test_dir = "test_edge_cases"
    os.makedirs(test_dir, exist_ok=True)

    # Test 1: Empty file
    print("\n1. Testing empty file...")
    with open(f"{test_dir}/empty.txt", 'w') as f:
        pass

    # Test 2: File with special characters in name
    print("2. Testing file with spaces and special chars...")
    with open(f"{test_dir}/file with spaces.txt", 'w') as f:
        f.write("Content with spaces in filename")

    # Test 3: Binary file
    print("3. Testing binary file...")
    with open(f"{test_dir}/binary.bin", 'wb') as f:
        f.write(bytes([0, 1, 2, 255, 128, 64, 32]))

    # Test 4: Unicode content
    print("4. Testing Unicode content...")
    with open(f"{test_dir}/unicode.txt", 'w', encoding='utf-8') as f:
        f.write("Hello 世界 🌍 Привет مرحبا")

    # Test 5: Large file
    print("5. Testing larger file...")
    with open(f"{test_dir}/large.txt", 'w') as f:
        for i in range(1000):
            f.write(f"Line {i}: " + "x" * 100 + "\n")

    # Test 6: Deep nested directory
    print("6. Testing deep nesting...")
    deep_dir = f"{test_dir}/a/b/c/d/e"
    os.makedirs(deep_dir, exist_ok=True)
    with open(f"{deep_dir}/deep.txt", 'w') as f:
        f.write("Deep nested file")

    # Pack all test files
    print("\n" + "=" * 80)
    print("PACKING...")
    print("=" * 80)
    packed_file = "test_edge_cases.txt"
    files_packed = pack_directory(test_dir, packed_file)
    print(f"\nPacked {files_packed} files")

    # Extract
    print("\n" + "=" * 80)
    print("EXTRACTING...")
    print("=" * 80)
    extract_dir = "test_edge_cases_extracted"
    files_extracted = extract_files_test(packed_file, extract_dir)

    # Verify all files
    print("\n" + "=" * 80)
    print("VERIFICATION...")
    print("=" * 80)

    test_files = [
        f"{test_dir}/empty.txt",
        f"{test_dir}/file with spaces.txt",
        f"{test_dir}/binary.bin",
        f"{test_dir}/unicode.txt",
        f"{test_dir}/large.txt",
        f"{deep_dir}/deep.txt"
    ]

    all_match = True
    for original in test_files:
        extracted = original.replace(test_dir, extract_dir, 1)

        if not os.path.exists(extracted):
            print(f"✗ MISSING: {extracted}")
            all_match = False
            continue

        with open(original, 'rb') as f1, open(extracted, 'rb') as f2:
            original_content = f1.read()
            extracted_content = f2.read()

            if original_content == extracted_content:
                print(f"✓ {os.path.basename(original)} ({len(original_content)} bytes)")
            else:
                print(f"✗ {original} MISMATCH!")
                print(f"  Original: {len(original_content)} bytes")
                print(f"  Extracted: {len(extracted_content)} bytes")
                all_match = False

    print("\n" + "=" * 80)
    if all_match:
        print("✓✓✓ ALL EDGE CASES PASSED!")
    else:
        print("✗✗✗ SOME TESTS FAILED!")
    print("=" * 80)

    return all_match

if __name__ == "__main__":
    success = test_edge_cases()
    sys.exit(0 if success else 1)
