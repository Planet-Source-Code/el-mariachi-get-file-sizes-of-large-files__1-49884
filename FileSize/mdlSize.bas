Attribute VB_Name = "mdlSize"
Option Explicit


Public Const GENERIC_READ = &H80000000
Public Const FILE_SHARE_READ = &H1
Public Const OPEN_EXISTING = 3
Public Const FILE_TYPE_CHAR = &H2
Public Const FILE_TYPE_DISK = &H1
Public Const FILE_TYPE_PIPE = &H3
Public Const FILE_TYPE_UNKNOWN = &H0
Public Declare Function GetFileType Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'**************************************
'Windows API/Global Declarations for :Fi
'     leFound()
'**************************************
Public Const MAX_PATH = 260
'Public Const MAXDWORD = &HFFFF
Public Const MAXDWORD As Long = &HFFFFFFFF
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Type FILETIME ' 8 Bytes
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type


Type WIN32_FIND_DATA ' 318 Bytes
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved_ As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Public Declare Function FindFirstFile& Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA)
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

'the traditional way to get the file size via the API
Function APIFileSize(strFileName As String) As Long
    Dim lpFindFileData As WIN32_FIND_DATA
    Dim hFindFirst As Long
    hFindFirst = FindFirstFile(strFileName, lpFindFileData)

    If hFindFirst > 0 Then
        FindClose hFindFirst
        APIFileSize = (lpFindFileData.nFileSizeHigh * MAXDWORD) + lpFindFileData.nFileSizeLow
    Else
        APIFileSize = 0
    End If
End Function

'the traditional revised way to get the file size via the API
Function APIFileSize2(strFileName As String) As Variant
    Dim lpFindFileData As WIN32_FIND_DATA
    Dim hFindFirst As Long
    hFindFirst = FindFirstFile(strFileName, lpFindFileData)

    If hFindFirst > 0 Then
        FindClose hFindFirst
        'APIFileSize2 = (lpFindFileData.nFileSizeHigh * MAXDWORD) + lpFindFileData.nFileSizeLow
        APIFileSize2 = FileLength(lpFindFileData.nFileSizeHigh, lpFindFileData.nFileSizeLow)
    Else
        APIFileSize2 = 0
    End If
End Function

Public Function FileLength(ByRef lFileSizeHigh As Long, ByRef lFileSizeLow As Long) As Variant
     Static Bit32 As Variant
     Dim FL As Variant
        
        
     If IsEmpty(Bit32) Then Bit32 = CDec(2 ^ 32)
     FL = CDec(0)
    
     If lFileSizeHigh < 0 Then
        FL = (Bit32 + CDec(lFileSizeHigh)) * Bit32
     Else
        FL = CDec(lFileSizeHigh) * Bit32
     End If
    
     If lFileSizeLow < 0 Then
        FileLength = FL + Bit32 + CDec(lFileSizeLow)
     Else
        FileLength = FL + CDec(lFileSizeLow)
     End If
End Function
        
Function MyFileSize(strFileName As String) As Variant
    Dim WFD As WIN32_FIND_DATA
    Dim hFindFirst As Long
    Dim Hi As Currency, Lo As Currency
    
    hFindFirst = FindFirstFile(strFileName, WFD)

    If hFindFirst > 0 Then
        FindClose hFindFirst
        Hi = WFD.nFileSizeHigh
        Lo = WFD.nFileSizeLow
        If Hi < 0 Then Hi = Hi + (2 ^ 32)
        If Lo < 0 Then Lo = Lo + (2 ^ 32)
        
        MyFileSize = (Hi * (2 ^ 32)) + Lo
    Else
        MyFileSize = 0
    End If
End Function

'this function uses the VB intrinsic function
Public Function VBFileSize(FilePath As String) As String
    Dim lngBSize As Long
    
    VBFileSize = ""
    DoEvents
    
    'the file size via FileLen
    lngBSize = FileLen(FilePath) 'size in bytes
    VBFileSize = Format(lngBSize, "#,##0")
End Function

'An alternarive API call that only works on NT based machines
Public Function WinNTSize(sFile As String) As Variant
    Dim hFile As Long, nSize As Currency, sSave As String
    
    nSize = 0
    
    'open the file
    hFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    'get the filesize
    GetFileSizeEx hFile, nSize
    'retrieve the file type
    Select Case GetFileType(hFile)
    Case FILE_TYPE_UNKNOWN
        sSave = "The type of the specified file is unknown"
    Case FILE_TYPE_DISK
        sSave = "The specified file is a disk file"
    Case FILE_TYPE_CHAR
        sSave = "The specified file is a character file, typically an LPT device or a console"
    Case FILE_TYPE_PIPE
        sSave = "The specified file is either a named or anonymous pipe"
    End Select
    'close the file
    CloseHandle hFile
    'MsgBox "File Type: " + sSave + vbCrLf + "Size:" + Str$(nSize * 10000) + " bytes"
    WinNTSize = nSize * 10000
End Function
