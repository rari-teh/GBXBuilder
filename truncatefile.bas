Attribute VB_Name = "countingpine"
'' see also https://stackoverflow.com/questions/6334917/how-can-i-trim-the-end-of-a-binary-file
'' based in part on http://vbnet.mvps.org/index.html?code/fileapi/truncaterandomfile.htm
Option Explicit

'' constants for CreateFile
Private Const OPEN_ALWAYS As Long = 4, GENERIC_WRITE As Long = &H40000000, GENERIC_READ As Long = &H80000000, FILE_ATTRIBUTE_NORMAL As Long = &H80, INVALID_HANDLE_VALUE As Long = -1

'' constants for SetFilePointer
Private Const FILE_BEGIN As Long = 0, INVALID_SET_FILE_POINTER As Long = -1

'' kernel32 functions needed
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hfile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hfile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hfile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hfile As Long) As Long

                    Sub truncatefile(filename As String, ByVal size As Double)
    Dim hfile As Long
    Dim dwFileSizeLow As Long
    Dim dwFileSizeHigh As Long
    Dim ret As Long

    '' open the file
    hfile = CreateFile(filename, _
        GENERIC_WRITE Or GENERIC_READ, _
        0&, ByVal 0&, _
        OPEN_ALWAYS, _
        FILE_ATTRIBUTE_NORMAL, _
        0&)

    Debug.Assert (hfile <> INVALID_HANDLE_VALUE) '' make sure file opened OK

    '' optional: get the current file length
    dwFileSizeLow = GetFileSize(hfile, dwFileSizeHigh)
    Debug.Assert (dwFileSizeLow >= 0 And dwFileSizeHigh = 0) '' TODO: handle 2GB and higher
    Debug.Print "Old file size: " & dwFileSizeLow

    '' split length into DWORDs (TODO: handle 2GB and higher)
    dwFileSizeLow = size
    dwFileSizeHigh = 0

    '' seek to the desired file length
    ret = SetFilePointer(hfile, dwFileSizeLow, dwFileSizeHigh, FILE_BEGIN)
    Debug.Assert ret <> INVALID_SET_FILE_POINTER

    '' set this as the length of the file
    ret = SetEndOfFile(hfile)
    Debug.Assert (ret <> 0)

    '' close the file handle
    Debug.Assert CloseHandle(hfile) <> 0

End Sub
