' Self-Extracted File Formatter - Extract Function
' Microsoft Access VBA Script
' Extracts files from the packed text file created by pack_files.py

Option Compare Database
Option Explicit

Public Sub ExtractFiles(packedFilePath As String, outputDir As String)
    '
    ' Extract all files from a packed text file to the output directory
    '
    ' Args:
    '   packedFilePath: Path to the packed text file
    '   outputDir: Directory where files will be extracted
    '
    Dim fso As Object
    Dim inputFile As Object
    Dim line As String
    Dim currentFilePath As String
    Dim currentFileSize As Long
    Dim currentEncoding As String
    Dim contentLines As String
    Dim inContent As Boolean
    Dim filesExtracted As Long

    On Error GoTo ErrorHandler

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check if packed file exists
    If Not fso.FileExists(packedFilePath) Then
        MsgBox "Packed file does not exist: " & packedFilePath, vbCritical
        Exit Sub
    End If

    ' Create output directory if it doesn't exist (with full path support)
    If Not fso.FolderExists(outputDir) Then
        Call CreateDirectoryPath(fso, outputDir)
    End If

    ' Open input file with UTF-8 encoding (TristateTrue = -1 for Unicode, but we need UTF-8)
    ' Using default encoding which will handle UTF-8 correctly
    Set inputFile = fso.OpenTextFile(packedFilePath, 1, False, -2) ' -2 = System default (UTF-8 compatible)

    inContent = False
    contentLines = ""
    filesExtracted = 0

    ' Parse the packed file
    Do Until inputFile.AtEndOfStream
        line = inputFile.ReadLine

        If InStr(line, "<<<FILE_START>>>") > 0 Then
            ' Reset for new file
            currentFilePath = ""
            currentFileSize = 0
            currentEncoding = ""
            contentLines = ""
            inContent = False

        ElseIf InStr(line, "PATH:") > 0 Then
            ' Extract file path
            currentFilePath = Trim(Mid(line, InStr(line, ":") + 1))

        ElseIf InStr(line, "SIZE:") > 0 Then
            ' Extract file size
            currentFileSize = CLng(Trim(Mid(line, InStr(line, ":") + 1)))

        ElseIf InStr(line, "ENCODING:") > 0 Then
            ' Extract encoding type
            currentEncoding = Trim(Mid(line, InStr(line, ":") + 1))

        ElseIf InStr(line, "<<<CONTENT_START>>>") > 0 Then
            ' Start collecting content
            inContent = True
            contentLines = ""

        ElseIf InStr(line, "<<<CONTENT_END>>>") > 0 Then
            ' Stop collecting content
            inContent = False

        ElseIf InStr(line, "<<<FILE_END>>>") > 0 Then
            ' Extract the file (allow empty files with contentLines = "")
            If currentFilePath <> "" Then
                Call WriteFile(fso, outputDir, currentFilePath, contentLines, currentEncoding)
                filesExtracted = filesExtracted + 1
                Debug.Print "Extracted: " & currentFilePath
            End If

        ElseIf inContent Then
            ' Collect content lines
            contentLines = contentLines & Trim(line)
        End If
    Loop

    inputFile.Close
    Set inputFile = Nothing
    Set fso = Nothing

    MsgBox "Successfully extracted " & filesExtracted & " files to " & outputDir, vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error extracting files: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    If Not inputFile Is Nothing Then
        inputFile.Close
        Set inputFile = Nothing
    End If
    Set fso = Nothing
End Sub

Private Sub WriteFile(fso As Object, baseDir As String, relativePath As String, encodedContent As String, encoding As String)
    '
    ' Write a file to disk with proper directory structure
    '
    ' Args:
    '   fso: FileSystemObject instance
    '   baseDir: Base output directory
    '   relativePath: Relative path for the file (using forward slashes)
    '   encodedContent: Encoded file content
    '   encoding: Encoding type (e.g., "base64")
    '
    Dim fullPath As String
    Dim dirPath As String
    Dim fileName As String
    Dim lastSlash As Integer
    Dim decodedBytes() As Byte
    Dim outputStream As Object

    ' Convert forward slashes to backslashes for Windows
    relativePath = Replace(relativePath, "/", "\")

    ' Build full path
    fullPath = fso.BuildPath(baseDir, relativePath)

    ' Extract directory path
    lastSlash = InStrRev(fullPath, "\")
    If lastSlash > 0 Then
        dirPath = Left(fullPath, lastSlash - 1)

        ' Create directory structure if it doesn't exist
        If Not fso.FolderExists(dirPath) Then
            Call CreateDirectoryPath(fso, dirPath)
        End If
    End If

    ' Decode content based on encoding type
    If LCase(encoding) = "base64" Then
        ' Handle empty files (no content)
        If Len(encodedContent) > 0 Then
            decodedBytes = Base64Decode(encodedContent)
        Else
            ' Empty file - create zero-length byte array
            ReDim decodedBytes(0)
        End If
    Else
        MsgBox "Unknown encoding: " & encoding, vbCritical
        Exit Sub
    End If

    ' Write binary file
    Set outputStream = CreateObject("ADODB.Stream")
    outputStream.Type = 1 ' Binary
    outputStream.Open

    ' Only write if there's content
    If Len(encodedContent) > 0 Then
        outputStream.Write decodedBytes
    End If

    outputStream.SaveToFile fullPath, 2 ' Overwrite if exists
    outputStream.Close
    Set outputStream = Nothing
End Sub

Private Sub CreateDirectoryPath(fso As Object, dirPath As String)
    '
    ' Recursively create directory path
    '
    Dim parentPath As String
    Dim lastSlash As Integer

    If fso.FolderExists(dirPath) Then
        Exit Sub
    End If

    ' Get parent directory
    lastSlash = InStrRev(dirPath, "\")
    If lastSlash > 0 Then
        parentPath = Left(dirPath, lastSlash - 1)
        If parentPath <> "" And Not fso.FolderExists(parentPath) Then
            Call CreateDirectoryPath(fso, parentPath)
        End If
    End If

    ' Create this directory
    fso.CreateFolder dirPath
End Sub

Private Function Base64Decode(base64String As String) As Byte()
    '
    ' Decode a base64 string to byte array
    '
    Dim xmlDoc As Object
    Dim node As Object

    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set node = xmlDoc.createElement("b64")

    node.DataType = "bin.base64"
    node.Text = base64String

    Base64Decode = node.nodeTypedValue

    Set node = Nothing
    Set xmlDoc = Nothing
End Function

' Example usage - can be called from Access form or module
Public Sub Example_ExtractPackedFiles()
    ' Change these paths as needed
    Dim packedFile As String
    Dim outputDirectory As String

    packedFile = "C:\temp\packed_files.txt"
    outputDirectory = "C:\temp\extracted"

    Call ExtractFiles(packedFile, outputDirectory)
End Sub
